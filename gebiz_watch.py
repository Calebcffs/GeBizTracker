#!/usr/bin/env python3
"""
gebiz_watch.py - Daily GeBiz opportunity scanner for SFSG.

Pulls GeBiz Business Opportunities RSS feeds (and, best-effort, public
listing pages), filters for container / modular / prefab / temporary-type
opportunities, dedupes against previously seen items, and posts a daily
digest to a Microsoft Teams channel via a Power Automate "Workflows"
webhook (the replacement for the retired Office 365 Incoming Webhook
connectors).

Pure Python standard library. No pip installs required.

Usage:
    python3 gebiz_watch.py                  # normal daily run
    python3 gebiz_watch.py --dry-run        # print report, don't post
    python3 gebiz_watch.py --selftest       # offline pipeline + card test

Webhook URL is read from env var TEAMS_WEBHOOK_URL, falling back to
"teams_webhook_url" in config.json.
"""

import argparse
import hashlib
import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

# ----------------------------------------------------------------------
# Time helpers (Singapore time, with fallback if tz database is missing)
# ----------------------------------------------------------------------
try:
    from zoneinfo import ZoneInfo
    SGT = ZoneInfo("Asia/Singapore")
except Exception:  # e.g. Windows without tzdata package
    SGT = timezone(timedelta(hours=8), name="SGT")


def now_sgt():
    return datetime.now(SGT)


# ----------------------------------------------------------------------
# Config / state
# ----------------------------------------------------------------------
DEFAULT_CONFIG = {
    "teams_webhook_url": "",
    "rss_feeds": [],           # [{"name": "...", "url": "https://..."}]
    "listing_pages": [],       # optional public GeBiz listing URLs (best-effort parse)
    "include_keywords": [
        "container", "modular", "prefab", "prefabricated", "ppvc",
        "portacabin", "portable cabin", "cabin", "site office",
        "temporary office", "temporary building", "temporary structure",
        "temporary classroom", "temporary quarters", "temporary shelter",
        "temporary holding", "temporary facility",
    ],
    "exclude_keywords": [
        "docker", "kubernetes", "containeris", "containeriz",
        "temporary staff", "temporary manpower", "temporary personnel",
    ],
    "closing_soon_days": 4,
    "max_new_items_in_card": 15,
    "max_closing_items_in_card": 10,
    "post_when_empty": True,
    "state_file": "seen_state.json",
    "request_timeout_seconds": 30,
    "user_agent": "Mozilla/5.0 (compatible; SFSG-GeBiz-Watch/1.0)",
}


def load_config(path):
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            user_cfg = json.load(f)
        for k, v in user_cfg.items():
            if not k.startswith("_"):
                cfg[k] = v
    else:
        log(f"WARNING: config file '{path}' not found, using built-in defaults")
    return cfg


def load_state(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            log(f"WARNING: could not read state file ({e}); starting fresh")
    return {}


def save_state(path, state):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=1, ensure_ascii=False)
    os.replace(tmp, path)


def log(msg):
    print(f"[gebiz-watch] {msg}", flush=True)


# ----------------------------------------------------------------------
# Fetching
# ----------------------------------------------------------------------
def http_get(url, cfg):
    req = urllib.request.Request(url, headers={
        "User-Agent": cfg["user_agent"],
        "Accept": "application/rss+xml, application/xml, text/xml, text/html, */*",
    })
    with urllib.request.urlopen(req, timeout=cfg["request_timeout_seconds"]) as resp:
        return resp.read().decode("utf-8", errors="replace")


# ----------------------------------------------------------------------
# Parsing helpers
# ----------------------------------------------------------------------
# GeBiz reference number styles seen in the wild, tried in order:
#   NEA000ETQ25000080 / BCA000ETT25000009   (classic ETQ/ETT)
#   MHASPF03000025207 / DEFNGPP7125100617   (letters + long digits)
#   JTC25T0077                              (short agency refs)
DOC_NO_PATTERNS = [
    re.compile(r"\b([A-Z0-9]{6}ET[TQ]\d{8})\b"),
    re.compile(r"\b([A-Z]{3,8}\d{9,14})\b"),
    re.compile(r"\b([A-Z]{2,5}\d{2}[A-Z]{1,3}\d{3,6})\b"),
]

CLOSING_DATE_RES = [
    re.compile(r"Clos(?:ing|e)[^0-9]{0,20}(\d{1,2}/\d{1,2}/\d{4})(?:[,\s]+(\d{1,2}:\d{2}\s*(?:[APap][Mm])?))?"),
    re.compile(r"Clos(?:ing|e)[^0-9]{0,20}(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})"),
]

DATE_FORMATS = ["%d/%m/%Y %I:%M %p", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%d %b %Y", "%d %B %Y", "%Y-%m-%d"]


def extract_doc_no(text):
    for pat in DOC_NO_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1)
    return None


def extract_closing(text):
    for pat in CLOSING_DATE_RES:
        m = pat.search(text)
        if m:
            date_part = m.group(1)
            time_part = (m.group(2) or "").strip() if m.lastindex and m.lastindex >= 2 else ""
            candidate = f"{date_part} {time_part}".strip()
            for fmt in DATE_FORMATS:
                try:
                    dt = datetime.strptime(candidate, fmt)
                    return dt.replace(tzinfo=SGT)
                except ValueError:
                    continue
            for fmt in DATE_FORMATS:
                try:
                    dt = datetime.strptime(date_part, fmt)
                    return dt.replace(tzinfo=SGT)
                except ValueError:
                    continue
    return None


def classify(doc_no, text):
    """Map GeBiz ref / wording to a procurement type with an S4-cap hint.

    Quotations (ETQ) are capped at S$90k by procurement rules, so they
    always fit an S4 (S$500k) registration. Tenders (ETT) exceed S$90k
    and GeBiz does not publish estimated values, so the S$500k ceiling
    is a manual check on those.
    """
    blob = f"{doc_no or ''} {text}".lower()
    if doc_no and "etq" in doc_no.lower() or "quotation" in blob:
        return "Quotation (<= ~S$90k, fits S4)"
    if doc_no and "ett" in doc_no.lower() or "tender" in blob:
        return "Tender (> S$90k, verify <= S$500k)"
    return "Type unknown"


def strip_tags(raw_html):
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", raw_html, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


def make_id(title, link):
    doc = extract_doc_no(f"{title} {link}")
    if doc:
        return doc
    return "H" + hashlib.sha1(f"{title}|{link}".encode("utf-8")).hexdigest()[:12].upper()


# ----------------------------------------------------------------------
# Source adapters -> normalized item dicts
# ----------------------------------------------------------------------
def parse_rss(xml_text, source_name):
    items = []
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        log(f"WARNING: RSS parse error for '{source_name}': {e}")
        return items
    for node in root.iter("item"):
        title = html.unescape((node.findtext("title") or "").strip())
        link = (node.findtext("link") or "").strip()
        desc = strip_tags(node.findtext("description") or "")
        if not title and not desc:
            continue
        blob = f"{title} {desc}"
        items.append({
            "id": make_id(title, link),
            "title": title or desc[:160],
            "detail": desc,
            "url": link or "https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml",
            "closing": extract_closing(blob),
            "source": source_name,
        })
    return items


def parse_listing_page(html_text, url, source_name):
    """Best-effort: pull GeBiz ref numbers + surrounding text off a public
    listing page. Crude, but gives partial coverage if no RSS URLs are
    configured yet. Items link back to the listing page itself."""
    text = strip_tags(html_text)
    items, seen_here = [], set()
    for pat in DOC_NO_PATTERNS:
        for m in pat.finditer(text):
            doc = m.group(1)
            if doc in seen_here:
                continue
            seen_here.add(doc)
            start = max(0, m.start() - 160)
            end = min(len(text), m.end() + 160)
            window = text[start:end].strip()
            items.append({
                "id": doc,
                "title": window[:200],
                "detail": window,
                "url": url,
                "closing": extract_closing(window),
                "source": source_name + " (page scan)",
            })
    return items


def collect_items(cfg):
    items, errors = [], []
    for feed in cfg.get("rss_feeds", []):
        name, url = feed.get("name", "RSS"), feed.get("url", "").strip()
        if not url or "PASTE" in url.upper():
            continue
        try:
            log(f"Fetching RSS: {name}")
            items.extend(parse_rss(http_get(url, cfg), name))
            time.sleep(1)
        except Exception as e:
            errors.append(f"RSS '{name}': {e}")
            log(f"ERROR fetching RSS '{name}': {e}")
    for url in cfg.get("listing_pages", []):
        url = url.strip()
        if not url:
            continue
        try:
            log(f"Fetching listing page: {url}")
            items.extend(parse_listing_page(http_get(url, cfg), url, "GeBiz listing"))
            time.sleep(1)
        except Exception as e:
            errors.append(f"Listing '{url}': {e}")
            log(f"ERROR fetching listing '{url}': {e}")
    # de-dupe across sources by id, keep first (RSS wins over page scan)
    unique = {}
    for it in items:
        unique.setdefault(it["id"], it)
    return list(unique.values()), errors


# ----------------------------------------------------------------------
# Filtering
# ----------------------------------------------------------------------
def match_keywords(item, cfg):
    blob = f"{item['title']} {item['detail']}".lower()
    for bad in cfg["exclude_keywords"]:
        if bad.lower() in blob:
            return None
    hits = [kw for kw in cfg["include_keywords"] if kw.lower() in blob]
    return hits or None


# ----------------------------------------------------------------------
# Report building
# ----------------------------------------------------------------------
def fmt_closing(dt):
    return dt.strftime("%a %d %b %Y") if dt else "closing date n/a"


def build_sections(matches, state, cfg):
    today = now_sgt().date()
    new_items, closing_soon = [], []

    for it in matches:
        if it["id"] not in state:
            new_items.append(it)

    horizon = today + timedelta(days=cfg["closing_soon_days"])
    new_ids = {it["id"] for it in new_items}
    for item_id, rec in state.items():
        if item_id in new_ids or not rec.get("closing"):
            continue
        try:
            cdate = datetime.fromisoformat(rec["closing"]).date()
        except ValueError:
            continue
        if today <= cdate <= horizon:
            closing_soon.append({
                "id": item_id,
                "title": rec.get("title", item_id),
                "url": rec.get("url", ""),
                "closing": datetime.fromisoformat(rec["closing"]),
                "kw": rec.get("kw", []),
            })
    closing_soon.sort(key=lambda x: x["closing"])
    return new_items, closing_soon


def update_state(state, matches):
    now_iso = now_sgt().isoformat()
    for it in matches:
        rec = state.get(it["id"], {"first_seen": now_iso})
        rec["title"] = it["title"][:250]
        rec["url"] = it["url"]
        rec["kw"] = it.get("kw", [])
        rec["closing"] = it["closing"].isoformat() if it["closing"] else rec.get("closing")
        state[it["id"]] = rec
    # prune: drop entries first seen > 180 days ago or closed > 30 days ago
    cutoff_seen = now_sgt() - timedelta(days=180)
    cutoff_closed = now_sgt() - timedelta(days=30)
    for k in list(state.keys()):
        rec = state[k]
        try:
            first = datetime.fromisoformat(rec.get("first_seen", now_iso))
            closed = datetime.fromisoformat(rec["closing"]) if rec.get("closing") else None
        except ValueError:
            continue
        if first < cutoff_seen or (closed and closed < cutoff_closed):
            del state[k]
    return state


def item_lines(it):
    kw = ", ".join(it.get("kw", []))
    doc_type = classify(it["id"], it["title"] + " " + it.get("detail", ""))
    meta = f"{it['id']} | {doc_type} | Closes: {fmt_closing(it.get('closing'))}"
    if kw:
        meta += f" | matched: {kw}"
    return it["title"], it["url"], meta


def build_text_report(new_items, closing_soon, scanned, errors, cfg):
    lines = [f"GeBiz Watch - {now_sgt().strftime('%a %d %b %Y, %H:%M')} SGT",
             f"Scanned {scanned} items | keywords: {', '.join(cfg['include_keywords'][:6])}..."]
    lines.append("")
    lines.append(f"NEW MATCHES ({len(new_items)})")
    if not new_items:
        lines.append("  none today")
    for it in new_items:
        title, url, meta = item_lines(it)
        lines += [f"  * {title}", f"    {meta}", f"    {url}"]
    lines.append("")
    lines.append(f"CLOSING WITHIN {cfg['closing_soon_days']} DAYS ({len(closing_soon)})")
    if not closing_soon:
        lines.append("  none")
    for it in closing_soon:
        lines.append(f"  * {it['title']} | closes {fmt_closing(it['closing'])} | {it['url']}")
    if errors:
        lines.append("")
        lines.append("FETCH ERRORS: " + "; ".join(errors))
    return "\n".join(lines)


# ----------------------------------------------------------------------
# Teams (Workflows webhook, Adaptive Card)
# ----------------------------------------------------------------------
def card_item_blocks(it):
    title, url, meta = item_lines(it)
    safe_title = title[:180].replace("]", ")").replace("[", "(")
    return [
        {"type": "TextBlock", "text": f"[{safe_title}]({url})" if url else safe_title,
         "wrap": True, "weight": "Bolder", "spacing": "Medium"},
        {"type": "TextBlock", "text": meta, "wrap": True,
         "isSubtle": True, "size": "Small", "spacing": "None"},
    ]


def build_adaptive_card(new_items, closing_soon, scanned, errors, cfg):
    body = [
        {"type": "TextBlock", "size": "Large", "weight": "Bolder",
         "text": f"GeBiz Watch - {now_sgt().strftime('%a %d %b %Y')}"},
        {"type": "TextBlock", "isSubtle": True, "wrap": True, "spacing": "None",
         "text": f"{len(new_items)} new match(es) | {len(closing_soon)} closing soon | {scanned} items scanned"},
    ]
    body.append({"type": "TextBlock", "text": f"New matches ({len(new_items)})",
                 "weight": "Bolder", "size": "Medium", "separator": True})
    if new_items:
        shown = new_items[:cfg["max_new_items_in_card"]]
        for it in shown:
            body.extend(card_item_blocks(it))
        if len(new_items) > len(shown):
            body.append({"type": "TextBlock", "isSubtle": True,
                         "text": f"...and {len(new_items) - len(shown)} more (see console log)"})
    else:
        body.append({"type": "TextBlock", "text": "None today.", "isSubtle": True})

    if closing_soon:
        body.append({"type": "TextBlock", "separator": True, "weight": "Bolder", "size": "Medium",
                     "text": f"Closing within {cfg['closing_soon_days']} days ({len(closing_soon)})"})
        for it in closing_soon[:cfg["max_closing_items_in_card"]]:
            body.extend(card_item_blocks(it))

    if errors:
        body.append({"type": "TextBlock", "separator": True, "wrap": True, "color": "Attention",
                     "text": "Fetch errors: " + "; ".join(errors)[:400]})

    body.append({"type": "TextBlock", "isSubtle": True, "size": "Small", "separator": True, "wrap": True,
                 "text": "ETQ quotations are <= ~S$90k (within S4). ETT tenders exceed S$90k; "
                         "confirm value fits the S$500k S4 cap before chasing."})

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "msteams": {"width": "Full"},
        "body": body,
    }
    return {"type": "message",
            "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive",
                             "contentUrl": None, "content": card}]}


def post_to_teams(payload, webhook_url, cfg):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(webhook_url, data=data, method="POST",
                                 headers={"Content-Type": "application/json"})
    for attempt in (1, 2):
        try:
            with urllib.request.urlopen(req, timeout=cfg["request_timeout_seconds"]) as resp:
                status = resp.getcode()
                if status in (200, 201, 202):
                    log(f"Posted to Teams (HTTP {status}). Note: 202 = accepted; "
                        "check the channel / workflow run history to confirm render.")
                    return True
                log(f"Unexpected HTTP {status} from webhook")
        except urllib.error.HTTPError as e:
            log(f"Webhook HTTP error {e.code}: {e.read().decode('utf-8', 'replace')[:300]}")
        except Exception as e:
            log(f"Webhook error: {e}")
        if attempt == 1:
            time.sleep(5)
    return False


# ----------------------------------------------------------------------
# Self-test
# ----------------------------------------------------------------------
SAMPLE_RSS = """<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"><channel><title>GeBIZ Business Opportunities - Test</title>
<item><title>NEA000ETQ26000123 - RENTAL OF OFFICE CONTAINERS AND MODULAR CABINS AT TUAS DEPOT</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>National Environment Agency | Published: 06/07/2026 | Closing Date: 20/07/2026 4:00 PM</description></item>
<item><title>MOE000ETT26000456 - CONSTRUCTION OF TEMPORARY CLASSROOM BUILDING (PREFABRICATED)</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>Ministry of Education | Closing Date: 09/07/2026 11:00 AM</description></item>
<item><title>GVT000ETQ26000789 - PROVISION OF TEMPORARY MANPOWER FOR EVENTS</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>Some Agency | Closing Date: 15/07/2026</description></item>
<item><title>IMD000ETT26000001 - KUBERNETES CONTAINER PLATFORM MAINTENANCE</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>IMDA | Closing Date: 30/07/2026</description></item>
<item><title>PUB000ETQ26000555 - SUPPLY OF WATER SAMPLING EQUIPMENT</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>PUB | Closing Date: 22/07/2026</description></item>
</channel></rss>"""


def selftest(cfg):
    log("SELFTEST: parsing sample RSS...")
    items = parse_rss(SAMPLE_RSS, "sample")
    assert len(items) == 5, f"expected 5 items, got {len(items)}"
    matches = []
    for it in items:
        hits = match_keywords(it, cfg)
        if hits:
            it["kw"] = hits
            matches.append(it)
    titles = " | ".join(m["id"] for m in matches)
    log(f"SELFTEST: {len(matches)} matched -> {titles}")
    assert len(matches) == 2, "keyword filter should keep 2 (containers/modular + temporary classroom)"
    assert all(m["closing"] for m in matches), "closing dates should parse"
    state = update_state({}, matches)
    new_items, closing_soon = build_sections(matches, {}, cfg)
    payload = build_adaptive_card(new_items, closing_soon, len(items), [], cfg)
    json.dumps(payload)  # must serialize
    size_kb = len(json.dumps(payload)) / 1024
    log(f"SELFTEST: adaptive card built OK ({size_kb:.1f} KB, Teams limit ~28 KB)")
    print(build_text_report(new_items, closing_soon, len(items), [], cfg))
    log("SELFTEST PASSED")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="GeBiz opportunity watcher -> Teams")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--state", default=None, help="override state file path")
    ap.add_argument("--dry-run", action="store_true", help="print report, do not post")
    ap.add_argument("--selftest", action="store_true", help="offline pipeline test")
    args = ap.parse_args()

    cfg = load_config(args.config)
    if args.selftest:
        selftest(cfg)
        return 0

    state_path = args.state or cfg["state_file"]
    state = load_state(state_path)

    items, errors = collect_items(cfg)
    if not items and errors:
        log("No items fetched and errors occurred; aborting without state change.")
    matches = []
    for it in items:
        hits = match_keywords(it, cfg)
        if hits:
            it["kw"] = hits
            matches.append(it)

    new_items, closing_soon = build_sections(matches, state, cfg)
    report = build_text_report(new_items, closing_soon, len(items), errors, cfg)
    print(report)

    if args.dry_run:
        log("Dry run: not posting, not saving state.")
        return 0

    posted = True
    if new_items or closing_soon or errors or cfg["post_when_empty"]:
        webhook = os.environ.get("TEAMS_WEBHOOK_URL", "").strip() or cfg.get("teams_webhook_url", "").strip()
        if not webhook or "PASTE" in webhook.upper():
            log("ERROR: no Teams webhook URL set (env TEAMS_WEBHOOK_URL or config). Skipping post.")
            posted = False
        else:
            payload = build_adaptive_card(new_items, closing_soon, len(items), errors, cfg)
            posted = post_to_teams(payload, webhook, cfg)
    else:
        log("Nothing to report and post_when_empty=false; skipping post.")

    if items:  # only advance state when we actually fetched something
        state = update_state(state, matches)
        save_state(state_path, state)
        log(f"State saved: {len(state)} tracked items -> {state_path}")

    return 0 if posted else 1


if __name__ == "__main__":
    sys.exit(main())
