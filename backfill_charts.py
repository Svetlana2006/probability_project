"""
backfill_charts.py
------------------
Fetch historical India Top 50 daily charts from charts.spotify.com for a
range of dates and append each as a new sheet in the workbook.

Uses the same persistent browser profile as append_daily_chart.py, so as
long as you have already logged in once you won't need to log in again.

Usage examples
--------------
# Fetch 15 Mar – 21 Mar 2026:
python backfill_charts.py --start 2026-03-15 --end 2026-03-21

# Fetch a single day:
python backfill_charts.py --start 2026-03-20 --end 2026-03-20

# Fetch April 7-8, 2026 (default if no dates provided):
python backfill_charts.py

# Use a different workbook:
python backfill_charts.py --start 2026-03-15 --end 2026-03-21 --input "My Workbook.xlsx"
"""

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

from spotify_pipeline import (
    resolve_browser_executable,
    resolve_workbook_path,
    fetch_india_top_50_browser_date,
    write_live_chart_sheet,
)


def date_range(start: date, end: date):
    """Yield each date from start to end inclusive."""
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Backfill historical India Top 50 Spotify charts into the workbook."
    )
    parser.add_argument("--start", help="Start date (YYYY-MM-DD). Defaults to 2026-04-07 if not provided.")
    parser.add_argument("--end", help="End date (YYYY-MM-DD). Defaults to 2026-04-08 if not provided.")
    parser.add_argument(
        "--input",
        default="Probability Project.xlsx",
        help="Path to the workbook.",
    )
    parser.add_argument(
        "--browser-profile",
        default=".browser-profile/spotify-charts",
        help="Persistent browser profile directory.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser headless (only if already logged in).",
    )
    args = parser.parse_args()

    # Default to backfilling April 20-22, 2026 when no explicit range is provided.
    start_str = args.start or "2026-04-24"
    end_str = args.end or "2026-04-25"

    try:
        start_date = date.fromisoformat(start_str)
        end_date = date.fromisoformat(end_str)
    except ValueError as exc:
        print(f"Invalid date format: {exc}")
        sys.exit(1)

    if start_date > end_date:
        print("--start must be on or before --end.")
        sys.exit(1)

    workbook_path = resolve_workbook_path(Path(args.input))
    browser_profile_dir = Path(args.browser_profile)
    browser_executable = resolve_browser_executable()

    dates = list(date_range(start_date, end_date))
    print(f"Fetching {len(dates)} chart(s): {start_str} → {end_str}\n")

    success, skipped, failed = 0, 0, 0

    for chart_date in dates:
        date_str = chart_date.isoformat()
        print(f"── {date_str} ", end="", flush=True)
        try:
            chart_df = fetch_india_top_50_browser_date(
                chart_date=date_str,
                browser_profile_dir=browser_profile_dir,
                browser_executable=browser_executable,
                headless=args.headless,
            )
            if chart_df is None or chart_df.empty:
                print("→ no data available (chart may not exist for this date)")
                skipped += 1
                continue
            sheet_name = write_live_chart_sheet(workbook_path, chart_df)
            print(f"→ saved as sheet {sheet_name}")
            success += 1
        except Exception as exc:
            print(f"→ FAILED: {exc}")
            failed += 1

    print(f"\nDone. {success} saved, {skipped} skipped, {failed} failed.")
    print(f"Workbook: {workbook_path}")


if __name__ == "__main__":
    main()
