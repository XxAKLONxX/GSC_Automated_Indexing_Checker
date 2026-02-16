# 🔍 GSC Automated Indexing Checker

<div align="center">

**Bulk-audit every URL on your site for Google Search Console indexing issues.**

Runs daily · Respects API quotas · Saves progress · Excel reports

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue?logo=python&logoColor=white)](https://python.org)
[![GSC API](https://img.shields.io/badge/API-Google%20Search%20Console-green?logo=google&logoColor=white)](https://developers.google.com/webmaster-tools)
[![MIT License](https://img.shields.io/badge/license-MIT-yellow)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)](#)

</div>

---

## ⚡ Quick Start

```bash
git clone https://github.com/YOUR_USERNAME/gsc-indexing-checker.git
cd gsc-indexing-checker
pip install -r requirements.txt
```

1. Add your `gsc-credentials.json` → [How to get it](#-google-service-account)
2. Edit `SITE_URL` in `gsc_auto_indexing_checker.py`
3. Run: `python gsc_auto_indexing_checker.py`
4. Schedule daily → sit back → open Excel when done

---

## 💡 Why This Exists

Google's URL Inspection API allows **2,000 requests/day**. If your site has 10k+ URLs, checking them manually is impossible.

This script handles it for you:

```
Day 1 → Fetch all URLs → Inspect first 2,000 → Save progress
Day 2 → Resume from 2,001 → Inspect next 2,000 → Update Excel
Day 3 → Resume from 4,001 → Inspect next 2,000 → Update Excel
Day 4 → Finish remaining → ✅ Complete report ready
```

> 💡 Each run takes ~20–30 min. Schedule it at 9 AM and forget about it.

---

## 📊 What You Get

### Excel Report (`GSC_Indexing_Report_MASTER.xlsx`)

| Sheet | What's Inside |
|:------|:-------------|
| 📋 **Summary** | Progress %, URLs checked/remaining, status, error count |
| 📄 **All URLs** | Every URL with full indexing details (15+ fields) |
| 📈 **Issue Summary** | Issue counts sorted by frequency |
| 🔍 **Per-Issue Sheets** | Filtered views for top 10 issue types |
| ⚠️ **Errors** | Any URLs that failed during inspection |

### Per-URL Data Collected

| Field | Field | Field |
|:------|:------|:------|
| Issue type | Verdict (PASS/FAIL) | Coverage state |
| Google canonical | User canonical | Canonical match |
| Last crawl time | Crawl agent | Robots.txt state |
| Page fetch state | Sitemap presence | Mobile usability |
| Rich results | Indexing state | Check timestamp |

---

## 🎯 Issues Detected

| Category | Issues |
|:---------|:-------|
| ✅ **Healthy** | Indexed (Submitted), Indexed (Valid) |
| 🟡 **Content** | Crawled not indexed, Discovered not indexed, Duplicate canonical conflicts |
| 🔴 **Errors** | 404, Soft 404, 5xx Server errors, 403 Forbidden |
| 🚫 **Blocked** | robots.txt, noindex, Redirects |
| ❓ **Other** | URL unknown to Google, Alternate with canonical, etc. |

---

## 🛠️ Setup

### Prerequisites

| Requirement | Details |
|:------------|:--------|
| Python | 3.7 or higher |
| Google Cloud | Project with Search Console API enabled |
| GSC Access | Service Account with **Owner** permission |

### 📦 Install

```bash
pip install -r requirements.txt

# Optional — Windows desktop notifications
pip install win10toast
```

### 🔑 Google Service Account

| Step | Action |
|:-----|:-------|
| 1 | Go to [Google Cloud Console](https://console.cloud.google.com/) → Create project |
| 2 | APIs & Services → Library → Enable **Google Search Console API** |
| 3 | IAM & Admin → Service Accounts → **Create Service Account** |
| 4 | Click account → Keys → **Add Key** → JSON → Download |
| 5 | Save as `gsc-credentials.json` in project folder |

### 🔗 Grant Access in GSC

| Step | Action |
|:-----|:-------|
| 1 | Open your JSON → copy the `client_email` value |
| 2 | [Search Console](https://search.google.com/search-console) → Settings → Users and permissions |
| 3 | **Add user** → paste email → set **Owner** permission |

> ⚠️ **Owner is required** for the URL Inspection API. Viewer/Full won't work.

### ⚙️ Configure

Edit `gsc_auto_indexing_checker.py`:

```python
SITE_URL = 'https://www.yoursite.com/'   # ← Must match GSC property exactly
```

> Include `https://` and trailing `/`

---

## 📅 Scheduling

### Windows → Task Scheduler

| Setting | Value |
|:--------|:------|
| Trigger | Daily at 9:00 AM |
| Action | Start program → `run_daily_checker.bat` |
| Start in | Your project folder path |
| Setting | ✅ Run task ASAP after missed start |

### Linux / macOS → cron

```bash
0 9 * * * cd /path/to/gsc-indexing-checker && python3 gsc_auto_indexing_checker.py >> cron.log 2>&1
```

---

## ⚙️ Configuration

All settings at the top of `gsc_auto_indexing_checker.py`:

| Variable | Default | Description |
|:---------|:--------|:-----------|
| `SITE_URL` | `https://www.example.com/` | Your GSC property URL |
| `DAILY_QUOTA` | `2000` | URLs per day (Google's limit) |
| `REQUEST_DELAY` | `0.5` | Seconds between requests |
| `OUTPUT_EXCEL` | `GSC_Indexing_Report_MASTER.xlsx` | Report filename |
| `PROGRESS_FILE` | `indexing_progress.json` | Progress tracker |
| `LOG_FILE` | `indexing_checker.log` | Execution log |

---

## 📁 Project Structure

```
gsc-indexing-checker/
├── gsc_auto_indexing_checker.py    ← Main script
├── emergency_recover.py            ← Crash recovery tool
├── run_daily_checker.bat           ← Windows batch runner
├── requirements.txt                ← Python dependencies
├── gsc-credentials.json.example    ← Credentials template
├── .gitignore                      ← Keeps secrets out of Git
├── LICENSE                         ← MIT
└── README.md
```

**Auto-generated at runtime** (git-ignored):

```
├── GSC_Indexing_Report_MASTER.xlsx ← 📊 Your report
├── indexing_progress.json          ← 💾 Progress state
├── indexing_checker.log            ← 📝 Execution log
└── run_log_*.txt                   ← 📝 Batch logs (Windows)
```

---

## ⏱️ Timeline

| Site Size | Days to Complete |
|:----------|:-----------------|
| 2,000 URLs | 1 day |
| 8,000 URLs | 4 days |
| 20,000 URLs | 10 days |
| 50,000 URLs | 25 days |
| 100,000+ URLs | 50+ days |

---

## 🚨 Constraints

| Constraint | Detail |
|:-----------|:-------|
| **2,000 URLs/day** | Hard limit by Google's URL Inspection API |
| **Quota resets** | Midnight Pacific Time |
| **Read-only** | Never modifies your GSC account |
| **Search Analytics only** | Orphan pages with 0 impressions won't appear |
| **Progress file** | Deleting `indexing_progress.json` mid-scan = restart from scratch |

---

## 🆘 Emergency Recovery

Script crashed mid-run? Progress not saved?

```bash
python emergency_recover.py
```

→ Parses log → finds checked URLs → fixes Excel → updates progress → next run continues normally.

---

## 🔧 Troubleshooting

| Problem | Solution |
|:--------|:---------|
| `Permission denied` / 403 | Add service account as **Owner** in GSC → Settings → Users |
| `Invalid credentials` | Re-download JSON key from Cloud Console |
| `Quota exceeded` / 429 | Wait until midnight PT — resumes automatically tomorrow |
| No notifications | `pip install win10toast` (optional, Windows only) |
| Script seems stuck | Check log file — each URL takes ~0.5s |
| Want fresh start | Delete `indexing_progress.json` → run again |
| Crash mid-run | Run `emergency_recover.py` → then run main script |

---

## 🔒 Security

- **Never commit** `gsc-credentials.json` — already in `.gitignore`
- Script uses **read-only** scope (`webmasters.readonly`)
- Revoke access anytime: GSC → Settings → Users → Remove

---

## 🤝 Contributing

Contributions welcome — open an issue or submit a PR.

## 📄 License

[MIT](LICENSE)
