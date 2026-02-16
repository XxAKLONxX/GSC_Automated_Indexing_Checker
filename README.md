# GSC Automated Indexing Checker

> Automatically audit every URL on your site for Google Search Console indexing issues. Runs daily, respects API quotas, saves progress, and produces a comprehensive Excel report.

---

## What It Does

Google's URL Inspection API has a hard limit of **2,000 requests per day**. If your site has 10,000+ URLs, manually checking them all is impractical.

This script automates the process: it inspects up to 2,000 URLs each day, saves its progress, and picks up exactly where it left off the next day. When finished, you get a single Excel workbook with every URL's indexing status, broken down by issue type.

### Issue Types Detected

The script captures every status the URL Inspection API returns, mapped to the same labels you see in Search Console:

**Healthy:** Indexed (Submitted and indexed), Indexed (Valid)

**Content issues:** Crawled — currently not indexed, Discovered — currently not indexed, Duplicate (Google chose different canonical), Duplicate without user-selected canonical, Alternate page with proper canonical tag

**Errors:** Not found (404), Soft 404, Server error (5xx), Access forbidden (403)

**Blocked:** Blocked by robots.txt, Excluded by noindex tag, Page with redirect

**Other:** URL unknown to Google, Submitted URL not found, and any other exclusion reason Google provides.

### Per-URL Data Collected

Each inspected URL produces a row with: issue type, verdict (PASS/FAIL/NEUTRAL), coverage state, indexing state, Google-selected canonical, user-declared canonical, canonical match flag, last crawl time, crawl agent (mobile/desktop), robots.txt state, page fetch state, sitemap presence, mobile usability verdict, and rich results status.

---

## How It Works

```
Day 1:  Fetch all URLs from Search Analytics → inspect first 2,000 → save to Excel + progress file
Day 2:  Load progress → inspect URLs 2,001–4,000 → update Excel
Day 3:  Load progress → inspect URLs 4,001–6,000 → update Excel
Day 4:  Load progress → inspect remaining URLs → mark complete
```

Each run takes roughly 20–30 minutes (0.5 s delay between requests × 2,000 URLs). Schedule it at 9 AM and forget about it.

---

## Requirements

- **Python 3.7+**
- A **Google Cloud project** with the Search Console API enabled
- A **Service Account** with Owner permission in your GSC property

### Python Dependencies

```
google-api-python-client >= 2.0.0
google-auth >= 2.0.0
google-auth-oauthlib >= 0.5.0
google-auth-httplib2 >= 0.1.0
pandas >= 1.3.0
openpyxl >= 3.0.9
xlsxwriter >= 3.0.0
```

Optional (Windows desktop notifications):
```
win10toast >= 0.9
```

---

## Setup

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/gsc-indexing-checker.git
cd gsc-indexing-checker
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt

# Optional: Windows notifications
pip install win10toast
```

### 3. Create a Google Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. **Enable the Google Search Console API** (APIs & Services → Library → search "Search Console API")
4. **Create a Service Account** (IAM & Admin → Service Accounts → Create)
5. **Generate a JSON key** (click the account → Keys tab → Add Key → JSON)
6. Save the downloaded file as `gsc-credentials.json` in the project folder

### 4. Grant Access in Search Console

1. Open the JSON file and copy the `client_email` value (looks like `name@project.iam.gserviceaccount.com`)
2. Go to [Google Search Console](https://search.google.com/search-console) → your property → Settings → Users and permissions
3. Click **Add user**, paste the service account email, set permission to **Owner**

> **Owner permission is required** for the URL Inspection API. Viewer/Full access is not enough.

### 5. Configure the Script

Open `gsc_auto_indexing_checker.py` and edit line 32:

```python
SITE_URL = 'https://www.yoursite.com/'  # Must match your GSC property exactly
```

Include the protocol (`https://`) and trailing slash. This must match your property URL in Search Console exactly.

### 6. Test

```bash
python gsc_auto_indexing_checker.py
```

You should see URLs being fetched and inspected. After the run, three files appear:

| File | Purpose |
|------|---------|
| `GSC_Indexing_Report_MASTER.xlsx` | Your main report (updated daily) |
| `indexing_progress.json` | Tracks which URLs have been checked — **do not delete mid-scan** |
| `indexing_checker.log` | Timestamped log of every run |

---

## Scheduling

### Windows (Task Scheduler)

1. Open Task Scheduler (`Win+R` → `taskschd.msc`)
2. Create Task (not "Basic Task")
3. **Trigger:** Daily at 9:00 AM
4. **Action:** Start a program → browse to `run_daily_checker.bat`; set "Start in" to your project folder
5. **Settings:** check "Run task as soon as possible after a scheduled start is missed"

### Linux / macOS (cron)

```bash
crontab -e

# Add this line (adjust paths):
0 9 * * * cd /path/to/gsc-indexing-checker && python3 gsc_auto_indexing_checker.py >> cron.log 2>&1
```

---

## Output: Excel Report

The report (`GSC_Indexing_Report_MASTER.xlsx`) contains these sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | Total URLs, checked count, remaining, progress %, status, error count |
| **All URLs** | Every inspected URL with all fields (issue type, verdict, canonicals, crawl info…) |
| **Issue Summary** | Count per issue type, sorted by frequency |
| **Per-issue sheets** | One sheet for each of the top 10 issue types (e.g. "Crawled - currently not i", "Not found (404)") — filtered for easy analysis |
| **Errors** | Any URLs that returned API errors during inspection |

---

## Project Structure

```
gsc-indexing-checker/
├── gsc_auto_indexing_checker.py   # Main script
├── emergency_recover.py           # Crash recovery tool
├── run_daily_checker.bat          # Windows batch runner
├── requirements.txt               # Python dependencies
├── gsc-credentials.json.example   # Template for credentials
├── LICENSE                        # MIT License
├── .gitignore                     # Keeps secrets out of Git
└── README.md                      # This file
```

**Generated at runtime (git-ignored):**

```
├── GSC_Indexing_Report_MASTER.xlsx   # Excel report
├── indexing_progress.json            # Progress state
├── indexing_checker.log              # Execution log
└── run_log_*.txt                     # Batch runner logs (Windows)
```

---

## Configuration

All settings are at the top of `gsc_auto_indexing_checker.py`:

| Variable | Default | Description |
|----------|---------|-------------|
| `SERVICE_ACCOUNT_FILE` | `gsc-credentials.json` | Path to your Google service account key |
| `SITE_URL` | `https://www.example.com/` | Your GSC property URL (must match exactly) |
| `DAILY_QUOTA` | `2000` | Max URLs to inspect per day (Google's limit) |
| `REQUEST_DELAY` | `0.5` | Seconds between API calls (avoid rate limiting) |
| `OUTPUT_EXCEL` | `GSC_Indexing_Report_MASTER.xlsx` | Report filename |
| `PROGRESS_FILE` | `indexing_progress.json` | Progress state filename |
| `LOG_FILE` | `indexing_checker.log` | Log filename |

---

## Constraints & Limitations

- **2,000 URLs/day** — this is a hard limit imposed by Google's URL Inspection API. A 10,000-URL site takes 5 days; a 100,000-URL site takes ~50 days.
- **Quota resets at midnight Pacific Time.**
- **Read-only** — the script never modifies anything in your GSC account. It only reads data.
- The script fetches URLs from **Search Analytics**, which means it only finds URLs that have received at least one impression. Orphan pages with zero impressions won't appear.
- The `win10toast` notification library is optional and Windows-only. The script works fine without it.
- If the `indexing_progress.json` file is deleted mid-scan, the script starts over from scratch.

---

## Emergency Recovery

If the script crashes mid-batch (power outage, network failure, 429 quota error) and fails to save progress:

```bash
python emergency_recover.py
```

This parses the log file to identify which URLs were successfully inspected, removes error rows from the Excel report, and updates the progress file so the next run continues from the correct position.

---

## Troubleshooting

| Problem | Cause | Solution |
|---------|-------|----------|
| "Permission denied" or 403 | Service account lacks Owner permission in GSC | Add the service account email as Owner in Search Console → Settings → Users |
| "Invalid credentials" | Missing or malformed JSON key | Re-download the key from Google Cloud Console |
| "Quota exceeded" / 429 | Already used 2,000 inspections today | Wait until midnight PT; the script will resume tomorrow |
| No Windows notifications | `win10toast` not installed | `pip install win10toast` (optional — doesn't affect core functionality) |
| Script appears stuck | Large site + slow connection | Check the log file for progress; each URL takes ~0.5 s |
| Want to restart from scratch | — | Delete `indexing_progress.json` and run again |
| Crash mid-run | Network issue, PC sleep, etc. | Run `python emergency_recover.py`, then run the main script again |

---

## Typical Timeline

| Site Size | Days to Complete |
|-----------|-----------------|
| 2,000 URLs | 1 day |
| 8,000 URLs | 4 days |
| 20,000 URLs | 10 days |
| 50,000 URLs | 25 days |
| 100,000 URLs | 50 days |

---

## Security Notes

- **Never commit `gsc-credentials.json`** — it's already in `.gitignore`.
- The service account only needs read access (the `webmasters.readonly` scope).
- You can revoke access at any time from Search Console → Settings → Users.

---

## Contributing

Contributions are welcome. Open an issue or submit a pull request.

---

## License

[MIT](LICENSE)
