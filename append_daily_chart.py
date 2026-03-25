import argparse
import sys
from pathlib import Path

from spotify_pipeline import (
    append_browser_chart_to_workbook,
    append_kworb_chart_to_workbook,
    resolve_browser_executable,
    resolve_workbook_path,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Append the latest India Top 50 daily chart as a new dated sheet in the workbook."
    )
    parser.add_argument(
        "--input",
        default="Probability Project.xlsx",
        help="Path to the workbook that stores the chart history.",
    )
    parser.add_argument(
        "--source",
        choices=["browser", "kworb"],
        default="browser",
        help="Import source. 'browser' uses the Spotify charts site with a persistent browser session; 'kworb' is a public fallback.",
    )
    parser.add_argument(
        "--browser-profile",
        default=".browser-profile/spotify-charts",
        help="Directory used to persist the browser login/session for Spotify charts.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run the browser importer headless. Leave this off for first-time login.",
    )
    parser.add_argument(
        "--login-timeout",
        type=int,
        default=300,
        help="Seconds to wait for you to complete Spotify login in the browser when needed.",
    )
    args = parser.parse_args()

    workbook_path = resolve_workbook_path(Path(args.input))
    try:
        if args.source == "browser":
            result = append_browser_chart_to_workbook(
                workbook_path=workbook_path,
                browser_profile_dir=Path(args.browser_profile),
                browser_executable=resolve_browser_executable(),
                headless=args.headless,
                login_timeout_seconds=args.login_timeout,
            )
            print(f"Browser profile: {result['browser_profile_dir']}")
        else:
            result = append_kworb_chart_to_workbook(workbook_path)
    except TimeoutError as exc:
        print(str(exc))
        print("Run the command again without --headless and complete the Spotify login in the opened browser window.")
        sys.exit(1)
    except Exception as exc:
        print(f"Append failed: {exc}")
        if args.source == "browser":
            print("If this is your first browser-auth run, use: python .\\append_daily_chart.py")
            print("That opens the browser profile so you can complete Spotify login once.")
            print("If Spotify browser auth is failing, you can fall back temporarily with: python .\\append_daily_chart.py --source kworb")
        sys.exit(1)

    print(f"Imported India Top 50 for {result['date']}")
    print(f"Workbook: {result['workbook_path']}")
    print(f"Sheet: {result['sheet_name']}")


if __name__ == "__main__":
    main()
