"""
Emergency Recovery Script
=========================
Recovers data when the main checker script crashes mid-run (e.g. after a
quota 429 error or an unexpected exception) and fails to save results.

What it does:
1. Parses indexing_checker.log to find all successfully inspected URLs
2. Removes API-error rows from the Excel report
3. Re-adds recovered URLs with proper field mappings
4. Updates indexing_progress.json with correct counts
5. Appends a completion summary to the log

Run once after a crash to restore lost progress:
    python emergency_recover.py
"""

import json
import re
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import os

# ============================================================================
# FILE PATHS — same defaults as the main script
# ============================================================================

LOG_FILE = "indexing_checker.log"
PROGRESS_FILE = "indexing_progress.json"
EXCEL_FILE = "GSC_Indexing_Report_MASTER.xlsx"

# ============================================================================
# LOG PARSER
# ============================================================================

def parse_log_for_checked_urls(log_path, target_date=None):
    """
    Parse the log file and extract all successfully checked URLs.

    Args:
        log_path:    Path to indexing_checker.log
        target_date: Optional 'YYYY-MM-DD' string.  If provided, only URLs
                     checked on that date are returned.  If None, the most
                     recent run date found in the log is used automatically.

    Returns:
        dict  {url: {'url', 'check_date', 'status_text'}}
    """
    print("[1/5] Parsing log file for checked URLs...")

    checked_urls = {}
    current_url = None
    current_timestamp = None
    dates_seen = set()

    with open(log_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()

        # Skip 429 / error lines — these URLs were NOT successfully checked
        if '[ERROR]' in line and ('429' in line or 'ERROR inspecting' in line):
            current_url = None
            continue

        # Match:  [YYYY-MM-DD HH:MM:SS] [INFO] [N/M] Checking: <URL>
        check_match = re.match(
            r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] \[INFO\] \[\d+/\d+\] Checking: (.+)',
            line
        )
        if check_match:
            timestamp_str, url = check_match.groups()
            date_part = timestamp_str[:10]
            dates_seen.add(date_part)

            if target_date is None or date_part == target_date:
                current_url = url.strip()
                current_timestamp = timestamp_str
            continue

        # Match:  [YYYY-MM-DD HH:MM:SS] [INFO]   ✓ <status text>
        if current_url and '[INFO]' in line and '✓' in line:
            status_match = re.search(r'✓\s+(.+)', line)
            if status_match:
                checked_urls[current_url] = {
                    'url': current_url,
                    'check_date': current_timestamp,
                    'status_text': status_match.group(1).strip()
                }
                current_url = None

    if not target_date and dates_seen:
        latest = max(dates_seen)
        print(f"   Auto-detected latest run date: {latest}")
        # Re-run filtered to that date
        return parse_log_for_checked_urls(log_path, target_date=latest)

    print(f"   ✓ Found {len(checked_urls)} successfully checked URLs"
          + (f" on {target_date}" if target_date else ""))
    return checked_urls


def map_status_to_fields(status_text):
    """
    Convert the human-readable status line from the log back into the
    structured fields used by the Excel report.
    """

    if "Submitted and indexed" in status_text:
        return {'issue_type': 'Indexed - Submitted and indexed',
                'verdict': 'PASS', 'coverage_state': 'Submitted and indexed',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Indexed" in status_text:
        return {'issue_type': 'Indexed - Valid',
                'verdict': 'PASS', 'coverage_state': 'Indexed',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Alternate page with proper canonical tag" in status_text:
        return {'issue_type': 'Alternate page with proper canonical tag',
                'verdict': 'NEUTRAL',
                'coverage_state': 'Alternate page with proper canonical tag',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Duplicate without user-selected canonical" in status_text:
        return {'issue_type': 'Duplicate without user-selected canonical',
                'verdict': 'NEUTRAL',
                'coverage_state': 'Duplicate without user-selected canonical',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Duplicate" in status_text:
        return {'issue_type': status_text, 'verdict': 'NEUTRAL',
                'coverage_state': status_text,
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Crawled - currently not indexed" in status_text:
        return {'issue_type': 'Crawled - currently not indexed',
                'verdict': 'NEUTRAL',
                'coverage_state': 'Crawled - currently not indexed',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "Discovered" in status_text:
        return {'issue_type': 'Discovered - currently not indexed',
                'verdict': 'NEUTRAL',
                'coverage_state': 'Discovered - currently not indexed',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "unknown to Google" in status_text or "URL is unknown" in status_text:
        return {'issue_type': 'Other - URL is unknown to Google',
                'verdict': 'FAIL',
                'coverage_state': 'URL is unknown to Google',
                'indexing_state': 'INDEXING_ALLOWED'}

    if "robots.txt" in status_text:
        return {'issue_type': 'Blocked by robots.txt', 'verdict': 'FAIL',
                'coverage_state': 'Blocked by robots.txt',
                'indexing_state': 'INDEXING_NOT_ALLOWED'}

    if "404" in status_text:
        return {'issue_type': 'Not found (404)', 'verdict': 'FAIL',
                'coverage_state': 'Not found (404)',
                'indexing_state': 'INDEXING_NOT_ALLOWED'}

    return {'issue_type': f'Other - {status_text}', 'verdict': 'NEUTRAL',
            'coverage_state': status_text,
            'indexing_state': 'INDEXING_ALLOWED'}


# ============================================================================
# EXCEL REPAIR
# ============================================================================

def update_excel_file(excel_path, checked_urls_dict):
    """Remove error rows and re-add recovered URLs to the 'All URLs' sheet."""
    print("\n[2/5] Updating Excel file...")

    df = pd.read_excel(excel_path, sheet_name='All URLs')
    original_count = len(df)
    print(f"   Original row count: {original_count}")

    df_clean = df[df['status'] != 'error'].copy()
    df_clean = df_clean[df_clean['issue_type'] != 'API Error'].copy()
    print(f"   ✓ Removed {original_count - len(df_clean)} error rows")

    existing_urls = set(df_clean['url'].tolist())

    new_rows = []
    for url, details in checked_urls_dict.items():
        if url not in existing_urls:
            fields = map_status_to_fields(details['status_text'])
            new_rows.append({
                'url': url,
                'issue_type': fields['issue_type'],
                'verdict': fields['verdict'],
                'coverage_state': fields['coverage_state'],
                'indexing_state': fields['indexing_state'],
                'google_canonical': '',
                'user_canonical': '',
                'canonical_match': 'No',
                'last_crawl_time': '',
                'crawled_as': 'MOBILE',
                'robots_txt_state': 'ALLOWED',
                'page_fetch_state': 'SUCCESSFUL',
                'in_sitemap': 'No',
                'mobile_friendly': 'VERDICT_UNSPECIFIED',
                'rich_results_status': 'N/A',
                'status': 'checked',
                'check_date': details['check_date'],
                'error_message': ''
            })

    print(f"   ✓ Adding {len(new_rows)} recovered URLs")

    df_final = pd.concat([df_clean, pd.DataFrame(new_rows)], ignore_index=True)

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a',
                        if_sheet_exists='replace') as writer:
        df_final.to_excel(writer, sheet_name='All URLs', index=False)

    print(f"   ✓ Excel updated: {len(df_final)} total rows")
    return len(df_final)


# ============================================================================
# PROGRESS JSON REPAIR
# ============================================================================

def update_progress_json(progress_path, checked_urls_dict):
    """Merge recovered URLs into the progress file."""
    print("\n[3/5] Updating progress JSON...")

    with open(progress_path, 'r', encoding='utf-8') as f:
        progress = json.load(f)

    print(f"   Current urls_checked: {progress.get('urls_checked', 0)}")

    existing_checked = {item['url'] for item in progress.get('results', [])}
    new_count = 0

    for url, details in checked_urls_dict.items():
        if url not in existing_checked:
            fields = map_status_to_fields(details['status_text'])
            progress['results'].append({
                'url': url,
                'issue_type': fields['issue_type'],
                'verdict': fields['verdict'],
                'coverage_state': fields['coverage_state'],
                'indexing_state': fields['indexing_state'],
                'google_canonical': '',
                'user_canonical': '',
                'canonical_match': 'No',
                'last_crawl_time': '',
                'crawled_as': 'MOBILE',
                'robots_txt_state': 'ALLOWED',
                'page_fetch_state': 'SUCCESSFUL',
                'in_sitemap': 'No',
                'mobile_friendly': 'VERDICT_UNSPECIFIED',
                'rich_results_status': 'N/A',
                'status': 'checked',
                'check_date': details['check_date']
            })
            new_count += 1

    progress['urls_checked'] = len(progress['results'])

    checked_url_set = {item['url'] for item in progress['results']}
    progress['urls_remaining'] = [
        u for u in progress['urls_remaining'] if u not in checked_url_set
    ]

    if 'errors' not in progress:
        progress['errors'] = []

    # Detect the date from recovered URLs
    if checked_urls_dict:
        sample = next(iter(checked_urls_dict.values()))
        progress['last_run_date'] = sample['check_date'][:10]

    with open(progress_path, 'w', encoding='utf-8') as f:
        json.dump(progress, f, indent=2, ensure_ascii=False)

    print(f"   ✓ Added {new_count} URLs to results")
    print(f"   ✓ New urls_checked: {progress['urls_checked']}")
    print(f"   ✓ URLs remaining: {len(progress['urls_remaining'])}")

    return progress


# ============================================================================
# LOG + SUMMARY REPAIR
# ============================================================================

def append_completion_to_log(log_path, progress, urls_added):
    """Write a synthetic completion block so the log looks consistent."""
    print("\n[4/5] Updating log file...")

    total_urls = len(progress.get('all_urls', []))
    pct = (progress['urls_checked'] / total_urls * 100) if total_urls else 0
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    lines = [
        f"[{ts}] [INFO] === EMERGENCY RECOVERY ===",
        f"[{ts}] [INFO] Recovered {urls_added} URLs from log",
        f"[{ts}] [INFO] Total progress: {progress['urls_checked']}/{total_urls} ({pct:.1f}%)",
        f"[{ts}] [INFO] Progress saved successfully",
        f"[{ts}] [INFO] {'=' * 70}",
        ""
    ]

    with open(log_path, 'a', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print("   ✓ Recovery summary appended to log")


def update_excel_summary(excel_path, progress):
    """Re-write the Summary sheet with corrected numbers."""
    print("\n[5/5] Updating Excel Summary sheet...")

    total_urls = len(progress.get('all_urls', []))
    checked = progress['urls_checked']
    remaining = len(progress['urls_remaining'])
    pct = (checked / total_urls * 100) if total_urls else 0

    summary = pd.DataFrame({
        'Metric': ['Total URLs', 'Checked', 'Remaining', 'Progress %',
                    'Last Run Date', 'Status', 'Errors'],
        'Value': [total_urls, checked, remaining, f"{pct:.1f}%",
                  progress['last_run_date'], progress['status'],
                  len(progress.get('errors', []))]
    })

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a',
                        if_sheet_exists='replace') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

    print("   ✓ Summary sheet updated")


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 70)
    print("EMERGENCY RECOVERY SCRIPT")
    print("=" * 70)
    print()

    # --- Pre-flight checks ---
    for path in (LOG_FILE, PROGRESS_FILE, EXCEL_FILE):
        if not os.path.exists(path):
            print(f"[ERROR] Required file not found: {path}")
            print("        Make sure you run this from the same directory as the checker.")
            return

    try:
        checked_urls_dict = parse_log_for_checked_urls(LOG_FILE)

        if not checked_urls_dict:
            print("\n[ERROR] No recoverable URLs found in the log.")
            return

        final_rows = update_excel_file(EXCEL_FILE, checked_urls_dict)
        progress = update_progress_json(PROGRESS_FILE, checked_urls_dict)

        # Append to log only once
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            if 'EMERGENCY RECOVERY' not in f.read():
                append_completion_to_log(LOG_FILE, progress, len(checked_urls_dict))
            else:
                print("\n[4/5] Log already has recovery entry — skipping")

        update_excel_summary(EXCEL_FILE, progress)

        total = len(progress.get('all_urls', []))
        print(f"\n{'=' * 70}")
        print("RECOVERY COMPLETE!")
        print(f"{'=' * 70}")
        print(f"✓ Recovered {len(checked_urls_dict)} URLs")
        print(f"✓ Total progress: {progress['urls_checked']}/{total}")
        print(f"✓ Excel rows: {final_rows}")
        print(f"✓ Last run date: {progress['last_run_date']}")
        print(f"{'=' * 70}")

    except Exception as e:
        print(f"\n[FATAL ERROR] {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
