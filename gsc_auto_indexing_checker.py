"""
GSC Automated Daily Indexing Status Checker
============================================
Automatically checks all website URLs for Google Search Console indexing issues.
Handles API quota limits (2,000/day), saves progress, and resumes where it left off.

Features:
- Checks 2,000 URLs per day (API quota limit)
- Saves progress to resume across multiple days
- Creates/updates a single Excel report with multiple sheets
- Detects ALL GSC indexing issue types
- Optional Windows desktop notifications
- Comprehensive error logging

Usage:
    python gsc_auto_indexing_checker.py

Schedule with Task Scheduler (Windows) or cron (Linux/Mac) for daily automation.
"""

from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
from datetime import datetime
import time
import json
import os
import sys
from pathlib import Path

# ============================================================================
# CONFIGURATION - UPDATE THESE FOR YOUR SITE
# ============================================================================

SERVICE_ACCOUNT_FILE = 'gsc-credentials.json'
SITE_URL = 'https://www.example.com/'  # <-- CHANGE THIS to your site URL

# Output file (updated daily with new data)
OUTPUT_EXCEL = 'GSC_Indexing_Report_MASTER.xlsx'

# Progress tracking file (DO NOT DELETE while a scan is in progress)
PROGRESS_FILE = 'indexing_progress.json'

# Daily quota limit (Google allows 2,000 URL inspections per day)
DAILY_QUOTA = 2000

# Delay between API requests in seconds (avoids rate limiting)
REQUEST_DELAY = 0.5

# Log file
LOG_FILE = 'indexing_checker.log'

# ============================================================================
# LOGGING
# ============================================================================

def log(message, level="INFO"):
    """Write to both console and log file."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_message = f"[{timestamp}] [{level}] {message}"

    print(log_message)

    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_message + '\n')
    except Exception as e:
        print(f"Warning: Could not write to log file: {e}")


def show_notification(title, message):
    """Show a Windows 10+ desktop notification (optional, non-blocking)."""
    try:
        from win10toast import ToastNotifier
        toaster = ToastNotifier()
        toaster.show_toast(title, message, duration=10, threaded=True)
    except Exception:
        log(f"NOTIFICATION: {title} - {message}")

# ============================================================================
# PROGRESS MANAGEMENT
# ============================================================================

def initialize_progress():
    """Return a fresh progress dictionary."""
    return {
        'last_run_date': None,
        'urls_checked': 0,
        'urls_remaining': [],
        'all_urls': [],
        'results': [],
        'errors': [],
        'status': 'not_started'
    }


def load_progress():
    """Load progress from a previous run, migrating missing keys if needed."""
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                progress = json.load(f)

            # Migrate: ensure every expected key exists
            defaults = initialize_progress()
            for key in defaults:
                if key not in progress:
                    progress[key] = defaults[key]
                    log(f"Migrated missing key '{key}' into progress file", "WARNING")

            return progress
        except Exception as e:
            log(f"Error loading progress file: {e}", "ERROR")
            return initialize_progress()
    else:
        return initialize_progress()


def save_progress(progress):
    """Persist progress to disk."""
    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(progress, f, indent=2)
        log("Progress saved successfully")
    except Exception as e:
        log(f"ERROR saving progress: {e}", "ERROR")

# ============================================================================
# GOOGLE SEARCH CONSOLE API
# ============================================================================

def initialize_service():
    """Authenticate and return a Search Console API service object."""
    log("Initializing Google Search Console API...")

    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        service = build('searchconsole', 'v1', credentials=credentials)
        log("✓ API initialized successfully")
        return service
    except Exception as e:
        log(f"ERROR initializing API: {e}", "ERROR")
        show_notification("GSC Error", f"Failed to initialize API: {e}")
        sys.exit(1)


def get_all_urls_from_search_analytics(service):
    """Fetch every URL that has appeared in Search Analytics (paginated)."""
    log("Fetching URLs from Search Console...")

    all_urls = []
    start_row = 0
    row_limit = 25000

    try:
        while True:
            response = service.searchanalytics().query(
                siteUrl=SITE_URL,
                body={
                    'startDate': '2024-01-01',
                    'endDate': datetime.now().strftime('%Y-%m-%d'),
                    'dimensions': ['page'],
                    'rowLimit': row_limit,
                    'startRow': start_row
                }
            ).execute()

            rows = response.get('rows', [])
            if not rows:
                break

            for row in rows:
                all_urls.append(row['keys'][0])

            log(f"  Fetched {len(all_urls)} URLs so far...")

            if len(rows) < row_limit:
                break

            start_row += row_limit
            time.sleep(1)

        log(f"✓ Total URLs found: {len(all_urls)}")
        return all_urls

    except Exception as e:
        log(f"ERROR fetching URLs: {e}", "ERROR")
        return []


def inspect_url(service, url):
    """
    Inspect a single URL via the URL Inspection API.

    Returns a dict with all indexing details (issue type, verdict,
    canonical info, crawl info, etc.) or an error record on failure.
    """
    try:
        request_body = {
            'inspectionUrl': url,
            'siteUrl': SITE_URL
        }

        response = service.urlInspection().index().inspect(body=request_body).execute()

        inspection = response.get('inspectionResult', {})
        index_status = inspection.get('indexStatusResult', {})
        mobile_usability = inspection.get('mobileUsabilityResult', {})
        rich_results = inspection.get('richResultsResult', {})

        verdict = index_status.get('verdict', 'UNKNOWN')
        coverage_state = index_status.get('coverageState', 'UNKNOWN')
        indexing_state = index_status.get('indexingState', 'UNKNOWN')
        robots_txt_state = index_status.get('robotsTxtState', 'UNKNOWN')
        page_fetch_state = index_status.get('pageFetchState', 'UNKNOWN')

        google_canonical = index_status.get('googleCanonical', '')
        user_canonical = index_status.get('userCanonical', '')

        last_crawl_time = index_status.get('lastCrawlTime', 'Never crawled')
        crawled_as = index_status.get('crawledAs', 'N/A')

        referring_urls = index_status.get('referringUrls', [])
        sitemap_urls = index_status.get('sitemap', [])

        issue_type = determine_issue_type(
            verdict, coverage_state, indexing_state,
            robots_txt_state, page_fetch_state,
            google_canonical, user_canonical
        )

        return {
            'url': url,
            'issue_type': issue_type,
            'verdict': verdict,
            'coverage_state': coverage_state,
            'indexing_state': indexing_state,
            'google_canonical': google_canonical,
            'user_canonical': user_canonical,
            'canonical_match': 'Yes' if google_canonical == user_canonical else 'No',
            'last_crawl_time': last_crawl_time,
            'crawled_as': crawled_as,
            'robots_txt_state': robots_txt_state,
            'page_fetch_state': page_fetch_state,
            'in_sitemap': 'Yes' if sitemap_urls else 'No',
            'mobile_friendly': mobile_usability.get('verdict', 'N/A'),
            'rich_results_status': rich_results.get('verdict', 'N/A'),
            'status': 'checked',
            'check_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    except Exception as e:
        error_msg = str(e)
        log(f"  ERROR inspecting {url}: {error_msg}", "ERROR")

        return {
            'url': url,
            'issue_type': 'API Error',
            'verdict': 'ERROR',
            'coverage_state': 'ERROR',
            'indexing_state': 'ERROR',
            'error_message': error_msg,
            'status': 'error',
            'check_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }


def determine_issue_type(verdict, coverage_state, indexing_state,
                         robots_txt_state, page_fetch_state,
                         google_canonical, user_canonical):
    """
    Map raw API fields to a human-readable issue label that matches
    the categories shown in the Search Console UI.
    """

    # Indexed
    if verdict == 'PASS':
        if coverage_state == 'Submitted and indexed':
            return 'Indexed - Submitted and indexed'
        return 'Indexed - Valid'

    # Duplicates
    if 'duplicate' in coverage_state.lower():
        if google_canonical != user_canonical and user_canonical:
            return 'Duplicate - Google chose different canonical than user'
        return 'Duplicate without user-selected canonical'

    # Alternate page with canonical
    if 'alternate' in coverage_state.lower():
        return 'Alternate page with proper canonical tag'

    # Crawled but not indexed
    if 'crawled' in coverage_state.lower() and 'not indexed' in coverage_state.lower():
        return 'Crawled - currently not indexed'

    # Discovered but not crawled
    if 'discovered' in coverage_state.lower():
        return 'Discovered - currently not indexed'

    # Blocked by robots.txt
    if robots_txt_state == 'BLOCKED':
        return 'Blocked by robots.txt'

    # Noindex
    if 'noindex' in coverage_state.lower():
        return 'Excluded by noindex tag'

    # 404 errors
    if page_fetch_state == 'NOT_FOUND' or '404' in coverage_state:
        if 'soft' in coverage_state.lower():
            return 'Soft 404'
        return 'Not found (404)'

    # Submitted URL not found
    if 'submitted' in coverage_state.lower() and '404' in coverage_state:
        return 'Submitted URL not found (404)'

    # Redirect
    if 'redirect' in coverage_state.lower():
        return 'Page with redirect'

    # Server errors
    if (page_fetch_state == 'SERVER_ERROR'
            or '5xx' in coverage_state
            or 'server error' in coverage_state.lower()):
        return 'Server error (5xx)'

    # Access forbidden
    if page_fetch_state == 'ACCESS_FORBIDDEN':
        return 'Access forbidden (403)'

    # Other exclusions
    if verdict == 'EXCLUDED':
        if indexing_state == 'INDEXING_NOT_ALLOWED':
            return f'Excluded - {coverage_state}'
        return 'Excluded by other reason'

    # Generic errors
    if verdict in ('ERROR', 'FAIL'):
        return f'Error - {coverage_state}'

    # Fallback
    return f'Other - {coverage_state}'

# ============================================================================
# DAILY BATCH PROCESSING
# ============================================================================

def process_daily_batch(service, progress):
    """Check up to DAILY_QUOTA URLs and update progress."""

    today = datetime.now().strftime('%Y-%m-%d')

    # Guard: already ran today
    if progress['last_run_date'] == today:
        log("Already ran today. Exiting.")
        show_notification("GSC Checker", "Already ran today. Next run tomorrow.")
        return progress

    # First run or previous scan finished — fetch fresh URL list
    if not progress['urls_remaining']:
        log("No URLs in queue. Fetching fresh URL list...")
        all_urls = get_all_urls_from_search_analytics(service)

        if not all_urls:
            log("ERROR: No URLs found!", "ERROR")
            show_notification("GSC Error", "No URLs found to check!")
            return progress

        progress['all_urls'] = all_urls
        progress['urls_remaining'] = all_urls.copy()
        progress['status'] = 'in_progress'

    urls_to_check = progress['urls_remaining'][:DAILY_QUOTA]
    total_urls = len(progress['all_urls'])
    urls_already_checked = progress['urls_checked']

    log("=" * 70)
    log(f"Starting daily batch for {today}")
    log(f"Total URLs: {total_urls}")
    log(f"Already checked: {urls_already_checked}")
    log(f"Remaining: {len(progress['urls_remaining'])}")
    log(f"Will check today: {len(urls_to_check)}")
    log("=" * 70)

    show_notification(
        "GSC Checker Started",
        f"Checking {len(urls_to_check)} URLs today\n"
        f"{urls_already_checked}/{total_urls} complete"
    )

    results_today = []
    errors_today = []

    for i, url in enumerate(urls_to_check, 1):
        try:
            log(f"[{i}/{len(urls_to_check)}] Checking: {url}")

            result = inspect_url(service, url)
            results_today.append(result)

            if result['status'] == 'error':
                errors_today.append(result)
                log(f"  ✗ ERROR: {result.get('error_message', 'Unknown error')}", "ERROR")
            else:
                log(f"  ✓ {result['issue_type']}")

            time.sleep(REQUEST_DELAY)

            # Progress notification every 100 URLs
            if i % 100 == 0:
                progress_pct = ((urls_already_checked + i) / total_urls) * 100
                log(f"Progress: {progress_pct:.1f}% complete")
                show_notification(
                    "Progress Update",
                    f"{i}/{len(urls_to_check)} checked today\n"
                    f"{progress_pct:.1f}% total progress"
                )

        except KeyboardInterrupt:
            log("Interrupted by user. Saving progress...", "WARNING")
            break
        except Exception as e:
            log(f"Unexpected error processing {url}: {e}", "ERROR")
            errors_today.append({
                'url': url,
                'error': str(e),
                'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })

    # Persist results
    progress['results'].extend(results_today)
    progress['errors'].extend(errors_today)
    progress['urls_checked'] += len(results_today)
    progress['urls_remaining'] = progress['urls_remaining'][len(results_today):]
    progress['last_run_date'] = today

    if not progress['urls_remaining']:
        progress['status'] = 'completed'
        log("=" * 70)
        log("ALL URLS CHECKED! Processing complete!")
        log("=" * 70)
        show_notification("GSC Checker COMPLETE!",
                          f"All {total_urls} URLs have been checked!")

    save_progress(progress)

    log(f"✓ Checked {len(results_today)} URLs today")
    log(f"✓ Total progress: {progress['urls_checked']}/{total_urls}")
    log(f"✗ Errors today: {len(errors_today)}")

    return progress

# ============================================================================
# EXCEL EXPORT
# ============================================================================

def export_to_excel(progress):
    """Write all results to a multi-sheet Excel workbook."""
    log("Exporting results to Excel...")

    if not progress['results']:
        log("No results to export yet.")
        return

    try:
        df_all = pd.DataFrame(progress['results'])

        with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:

            # --- Summary sheet ---
            total_urls = len(progress['all_urls'])
            checked_urls = progress['urls_checked']
            remaining_urls = len(progress['urls_remaining'])

            summary_data = {
                'Metric': [
                    'Total URLs', 'Checked', 'Remaining',
                    'Progress %', 'Last Run Date', 'Status', 'Errors'
                ],
                'Value': [
                    total_urls, checked_urls, remaining_urls,
                    f"{(checked_urls / total_urls * 100):.1f}%" if total_urls else "0%",
                    progress['last_run_date'] or 'Never',
                    progress['status'],
                    len(progress['errors'])
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

            # --- All URLs sheet ---
            df_all.to_excel(writer, sheet_name='All URLs', index=False)

            # --- Issue summary + per-issue sheets ---
            if 'issue_type' in df_all.columns:
                issue_counts = df_all['issue_type'].value_counts()

                pd.DataFrame({
                    'Issue Type': issue_counts.index,
                    'Count': issue_counts.values
                }).to_excel(writer, sheet_name='Issue Summary', index=False)

                for issue_type in issue_counts.head(10).index:
                    df_issue = df_all[df_all['issue_type'] == issue_type]
                    sheet_name = issue_type[:31]  # Excel sheet name limit
                    df_issue.to_excel(writer, sheet_name=sheet_name, index=False)

            # --- Errors sheet ---
            if progress['errors']:
                pd.DataFrame(progress['errors']).to_excel(
                    writer, sheet_name='Errors', index=False
                )

        log(f"✓ Excel file updated: {OUTPUT_EXCEL}")
        log(f"  Sheets: Summary, All URLs, Issue Summary, + individual issue sheets")

    except Exception as e:
        log(f"ERROR creating Excel file: {e}", "ERROR")

# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    """Run one daily batch: authenticate, check URLs, export Excel."""

    start_time = datetime.now()
    log("=" * 70)
    log("GSC AUTOMATED INDEXING STATUS CHECKER")
    log("=" * 70)
    log(f"Started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Site: {SITE_URL}")
    log("")

    try:
        service = initialize_service()

        progress = load_progress()
        log(f"Current status: {progress['status']}")

        progress = process_daily_batch(service, progress)

        export_to_excel(progress)

        end_time = datetime.now()
        duration = (end_time - start_time).seconds
        log("")
        log("=" * 70)
        log(f"Completed at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        log(f"Time taken: {duration} seconds")
        log("=" * 70)

        if progress['status'] == 'completed':
            show_notification("GSC Checker - ALL DONE!",
                              f"All URLs checked!\nResults in {OUTPUT_EXCEL}")
        else:
            remaining = len(progress['urls_remaining'])
            show_notification("GSC Checker - Done for Today",
                              f"Next run tomorrow\n{remaining} URLs remaining")

    except Exception as e:
        log(f"FATAL ERROR: {e}", "ERROR")
        show_notification("GSC Checker ERROR",
                          f"Fatal error occurred:\n{str(e)[:100]}")
        raise


if __name__ == '__main__':
    main()
