#!/usr/bin/env python3
"""
gebiz_watch.py - Daily GeBiz opportunity scanner for SFSG.

Pulls GeBiz Business Opportunities listing pages (and RSS feeds if
configured), filters for container / modular / prefab / temporary-type
opportunities, dedupes against previously seen items, and posts a daily
digest to a Microsoft Teams channel via a Power Automate "Workflows"
webhook (the replacement for the retired Office 365 Incoming Webhook
connectors).

Listing pages are fetched via a headless Chromium browser (Playwright)
to pass Imperva WAF. Install once:
    pip install playwright
    playwright install chromium

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


_GEBIZ_SEARCH_URL = "https://www.gebiz.gov.sg/ptn/opportunity/BOAdvancedSearch.xhtml"


def playwright_get(url, cfg):
    """Fetch GeBiz listing via headless Chromium (bypasses Imperva WAF).

    Runs one GeBiz Advanced Search per include_keyword so we get all open
    tenders matching any of our terms without needing to paginate the full
    listing (which loops back to page 1 due to JSF ViewState issues).
    The url argument is accepted for interface compatibility but not used.
    """
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        raise RuntimeError(
            "playwright is required for listing pages.\n"
            "Run: pip install playwright && playwright install chromium"
        )
    timeout_ms = cfg.get("request_timeout_seconds", 30) * 1000
    keywords = cfg.get("include_keywords", [])
    all_html = []
    seen_docs = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ))

        for kw in keywords:
            # Navigate to Advanced Search (fresh form + Imperva session)
            page.goto(_GEBIZ_SEARCH_URL, wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                page.wait_for_selector("input[type='submit'][value='Search']", timeout=15000)
            except PWTimeout:
                log(f"  WARNING: search form not ready for '{kw}', skipping")
                continue

            # Fill the keyword/description text field (first visible text input on the form)
            filled = page.evaluate(f"""
                () => {{
                    const field = [...document.querySelectorAll('input[type="text"]')]
                        .find(el => el.offsetParent !== null);
                    if (field) {{ field.value = {json.dumps(kw)}; return true; }}
                    return false;
                }}
            """)
            if not filled:
                log(f"  WARNING: could not find search input field for '{kw}'")
                continue

            btn = page.query_selector("input[type='submit'][value='Search']")
            if btn:
                btn.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=timeout_ms)
                except PWTimeout:
                    pass

            html_chunk = page.content()
            new_docs = set(re.findall(r'\b[A-Z0-9]{6}ET[TQ]\d{8}\b', html_chunk)) - seen_docs
            seen_docs.update(new_docs)
            all_html.append(html_chunk)
            log(f"  '{kw}': {len(new_docs)} new doc(s)")

        browser.close()
    return "\n".join(all_html)


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


def parse_gebiz_row(text):
    """Extract clean title and agency from a GeBiz listing row text window.

    GeBiz rows follow the pattern:
        ... OPEN {title} LOADING Agency {agency} Published {date} ...
    Returns (title, agency) strings, or (None, None) if not found.
    """
    title = None
    agency = None
    m = re.search(r'\bOPEN\s+(.+?)\s+LOADING', text, re.I)
    if m:
        title = re.sub(r'\s+', ' ', m.group(1)).strip()
    m = re.search(r'\bAgency\s+(.+?)(?:\s*\||\s+Published|\s+Clos)', text, re.I)
    if m:
        agency = re.sub(r'[|\s]+$', '', re.sub(r'\s+', ' ', m.group(1))).strip()
    return title, agency


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
        _, agency = parse_gebiz_row(f"{title} LOADING {desc}")
        items.append({
            "id": make_id(title, link),
            "title": title or desc[:160],
            "agency": agency or "",
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
            end = min(len(text), m.end() + 320)
            window = text[start:end].strip()
            clean_title, agency = parse_gebiz_row(window)
            items.append({
                "id": doc,
                "title": clean_title or window[:200],
                "agency": agency or "",
                "detail": window,
                "url": f"https://www.gebiz.gov.sg/ptn/opportunity/directlink.xhtml?docCode={doc}",
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
            items.extend(parse_listing_page(playwright_get(url, cfg), url, "GeBiz listing"))
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
    """Return (all_matches, new_ids) where new_ids is the set of IDs not yet in state."""
    new_ids = {it["id"] for it in matches if it["id"] not in state}
    return matches, new_ids


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
    agency = it.get("agency", "")
    meta = f"{it['id']} | {doc_type} | Closes: {fmt_closing(it.get('closing'))}"
    if agency:
        meta = f"{agency} | " + meta
    if kw:
        meta += f" | matched: {kw}"
    return it["title"], it["url"], meta


def build_text_report(matches, new_ids, scanned, errors, cfg):
    new_count = sum(1 for it in matches if it["id"] in new_ids)
    lines = [f"GeBiz Watch - {now_sgt().strftime('%a %d %b %Y, %H:%M')} SGT",
             f"Scanned {scanned} items | keywords: {', '.join(cfg['include_keywords'][:6])}..."]
    lines.append("")
    lines.append(f"OPEN MATCHES ({len(matches)}, {new_count} new today)")
    if not matches:
        lines.append("  none")
    for it in matches:
        title, url, meta = item_lines(it)
        marker = "  *[NEW] " if it["id"] in new_ids else "  * "
        lines += [f"{marker}{title}", f"    {meta}", f"    {url}"]
    if errors:
        lines.append("")
        lines.append("FETCH ERRORS: " + "; ".join(errors))
    return "\n".join(lines)


# ----------------------------------------------------------------------
# Teams (Workflows webhook, Adaptive Card)
# ----------------------------------------------------------------------
def card_item_blocks(it, is_new=False):
    title, url, meta = item_lines(it)
    safe_title = title[:180].replace("]", ")").replace("[", "(")
    display = f"[{safe_title}]({url})" if url else safe_title
    if is_new:
        display = "NEW: " + display
    return [
        {"type": "TextBlock", "text": display,
         "wrap": True, "weight": "Bolder", "spacing": "Medium"},
        {"type": "TextBlock", "text": meta, "wrap": True,
         "isSubtle": True, "size": "Small", "spacing": "None"},
    ]


def build_adaptive_card(matches, new_ids, scanned, errors, cfg):
    new_count = sum(1 for it in matches if it["id"] in new_ids)
    body = [
        {"type": "TextBlock", "size": "Large", "weight": "Bolder",
         "text": f"GeBiz Watch - {now_sgt().strftime('%a %d %b %Y')}"},
        {"type": "TextBlock", "isSubtle": True, "wrap": True, "spacing": "None",
         "text": f"{len(matches)} open match(es) ({new_count} new today) | {scanned} items scanned"},
    ]
    body.append({"type": "TextBlock", "text": f"Open matches ({len(matches)})",
                 "weight": "Bolder", "size": "Medium", "separator": True})
    if matches:
        shown = matches[:cfg["max_new_items_in_card"]]
        for it in shown:
            body.extend(card_item_blocks(it, is_new=it["id"] in new_ids))
        if len(matches) > len(shown):
            body.append({"type": "TextBlock", "isSubtle": True,
                         "text": f"...and {len(matches) - len(shown)} more (see console log)"})
    else:
        body.append({"type": "TextBlock", "text": "No open tenders matching keywords.", "isSubtle": True})

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
<description>Agency National Environment Agency | Published: 06/07/2026 | Closing Date: 20/07/2026 4:00 PM</description></item>
<item><title>MOE000ETT26000456 - CONSTRUCTION OF TEMPORARY CLASSROOM BUILDING (PREFABRICATED)</title>
<link>https://www.gebiz.gov.sg/ptn/opportunity/index.xhtml</link>
<description>Agency Ministry of Education | Closing Date: 09/07/2026 11:00 AM</description></item>
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
    all_matches, new_ids = build_sections(matches, {}, cfg)
    payload = build_adaptive_card(all_matches, new_ids, len(items), [], cfg)
    json.dumps(payload)  # must serialize
    size_kb = len(json.dumps(payload)) / 1024
    log(f"SELFTEST: adaptive card built OK ({size_kb:.1f} KB, Teams limit ~28 KB)")
    print(build_text_report(all_matches, new_ids, len(items), [], cfg))
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

    all_matches, new_ids = build_sections(matches, state, cfg)
    report = build_text_report(all_matches, new_ids, len(items), errors, cfg)
    print(report)

    if args.dry_run:
        log("Dry run: not posting, not saving state.")
        return 0

    posted = True
    if all_matches or errors or cfg["post_when_empty"]:
        webhook = os.environ.get("TEAMS_WEBHOOK_URL", "").strip() or cfg.get("teams_webhook_url", "").strip()
        if not webhook or "PASTE" in webhook.upper():
            log("ERROR: no Teams webhook URL set (env TEAMS_WEBHOOK_URL or config). Skipping post.")
            posted = False
        else:
            payload = build_adaptive_card(all_matches, new_ids, len(items), errors, cfg)
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
