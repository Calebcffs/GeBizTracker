# GeBiz Watch

Daily scanner for GeBiz opportunities relevant to SFSG (container, modular, prefab, temporary structures), posted as a digest card into a Microsoft Teams channel.

Flow: `GeBiz RSS feeds -> keyword filter -> dedupe against seen_state.json -> Adaptive Card -> Teams channel`

Pure Python 3.9+ standard library. Nothing to pip install.

---

## 1. Why the previous bot probably failed (read this first)

Microsoft **retired the old "Incoming Webhook" connectors** in Teams (Office 365 Connectors, fully decommissioned end-2025). Every older tutorial that says "Channel > Connectors > Incoming Webhook" now produces a dead or dying URL. The current method is a **Teams Workflow** (Power Automate under the hood), and it expects an **Adaptive Card** payload, not the old MessageCard format. This script already sends the correct format.

## 2. Create the Teams webhook (5 minutes, once)

1. In Teams, go to the target channel, click the **...** next to the channel name, choose **Workflows**.
2. Pick the template **"Post to a channel when a webhook request is received"**.
3. Confirm the Team and Channel, click **Add workflow**, and **copy the URL** it gives you (long `...powerplatform.com/...` or `...logic.azure.com/...` link).
4. Treat that URL like a password. Anyone with it can post to the channel.

Notes:
- Posts appear as "*YourName* via Workflows". If that bothers people, create the workflow from a shared/service account.
- A successful post returns HTTP **202** (accepted). If nothing appears in the channel, open the Workflows app > your flow > **Run history** to see the error.

Quick test from a terminal (should drop a card in the channel):

```bash
export TEAMS_WEBHOOK_URL="paste_url_here"
python3 - <<'EOF'
import json, os, urllib.request
card = {"type":"message","attachments":[{"contentType":"application/vnd.microsoft.card.adaptive",
  "content":{"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","type":"AdaptiveCard",
  "version":"1.4","body":[{"type":"TextBlock","text":"GeBiz Watch webhook test OK","weight":"Bolder"}]}}]}
req = urllib.request.Request(os.environ["TEAMS_WEBHOOK_URL"], data=json.dumps(card).encode(),
  headers={"Content-Type":"application/json"}, method="POST")
print(urllib.request.urlopen(req).getcode())
EOF
```

## 3. Get GeBiz RSS feed URLs (10 minutes, once)

GeBiz discontinued its email alerts (GoBusiness GeBiz Alerts ended Feb 2025). What remains official and free is **RSS feeds per procurement category**, available to registered GeBiz Trading Partners (GTP registration is free and separate from the financial-grade supplier registration).

1. Log in to GeBiz (Corppass) as a Trading Partner.
2. Go to **Opportunities** search and browse by **procurement category**.
3. Each category listing exposes an **RSS icon/link**. Copy the feed URL for every category worth watching. Suggested starting set: Construction and related works, Facilities Management, Logistics/Transport, and any category where past folder items (NEA, MHA, DEF, JTC ones) were filed.
4. Paste each URL into `config.json` under `rss_feeds`.

The feeds are per category, so the keyword filter in this script does the cross-category narrowing to container/modular/prefab/temporary.

If RSS is not set up yet, you can temporarily add public GeBiz listing URLs to `listing_pages` in `config.json`. That mode does a rough text scan of the page and is less reliable; treat it as a stopgap.

## 4. Run it

```bash
python3 gebiz_watch.py --selftest      # offline check: parsing, filtering, card building
python3 gebiz_watch.py --dry-run       # real fetch, prints report, posts nothing
python3 gebiz_watch.py                 # real run: fetch, post to Teams, save state
```

`seen_state.json` records every match already reported, so the daily card only shows genuinely new items, plus a "closing within 4 days" reminder section.

## 5. Schedule it

**Option A: GitHub Actions (recommended, laptop can be off)**
1. Push this folder to a **private** GitHub repo.
2. Repo Settings > Secrets and variables > Actions > New secret: `TEAMS_WEBHOOK_URL`.
3. The included workflow (`.github/workflows/gebiz-daily.yml`) runs weekdays at 08:30 SGT and commits `seen_state.json` back so the dedupe memory persists. Use the **Run workflow** button in the Actions tab for a first manual test.

**Option B: Windows Task Scheduler (simplest, laptop must be on)**
```
Program:  python
Arguments: C:\path\to\gebiz_watch.py --config C:\path\to\config.json
Trigger:  Daily 08:30
```
Set the webhook once with `setx TEAMS_WEBHOOK_URL "paste_url_here"` (new terminals only), or paste it into `config.json`.

## 6. The S$500k question

GeBiz does **not publish estimated contract values** on open notices, so a hard <= S$500k filter is impossible from the data. The script uses the procurement-rule proxy instead:

- **ETQ (quotation)** references: capped around S$90k by procurement rules, so always within the S4 grade. Auto-labelled "fits S4".
- **ETT (tender)** references: above S$90k, could be anything. Labelled "verify <= S$500k" so a human sanity-checks before chasing.

## 7. Tuning

- Add or remove terms in `include_keywords` and `exclude_keywords` (both are plain substrings, case-insensitive). The excludes stop the two classic false-positive floods: IT "container/Kubernetes" tenders and "temporary staff" tenders.
- `post_when_empty: true` posts a short "none today" card so silence means broken, not quiet. Set false to only post when there is something.
- `closing_soon_days` controls the reminder window.

## 8. Optional extra: competitor intel

data.gov.sg hosts MOF's "Government Procurement via GeBIZ" dataset (dataset id `d_acde1106003906a75c3fa052592f2fcb`). It is **awarded** contracts (supplier + awarded amount), updated periodically, so it cannot power daily alerts, but it is great for checking who won past container/modular jobs and at what price when preparing bids.
