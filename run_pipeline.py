import os
import argparse
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
MPL_CONFIG_DIR = PROJECT_DIR / ".matplotlib"
MPL_CONFIG_DIR.mkdir(exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(MPL_CONFIG_DIR))
os.environ.setdefault("MPLBACKEND", "Agg")

from spotify_pipeline import run_pipeline


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the full Spotify Top 50 probability pipeline from workbook to report."
    )
    parser.add_argument(
        "--input",
        default="Probability Project.xlsx",
        help="Path to the source Excel workbook.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Directory where tables, figures, and the report will be saved.",
    )
    parser.add_argument(
        "--cleaned-csv",
        default="spotify_data.csv",
        help="Path to the cleaned CSV output.",
    )
    parser.add_argument(
        "--metadata",
        default="genre_mapping.csv",
        help="Optional CSV file with Artist, Song, Genre columns for genre analysis.",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=2026,
        help="Year used when workbook sheet names are in DDMM format.",
    )
    parser.add_argument(
        "--live-fetch",
        action="store_true",
        help="Attempt the built-in Spotify live fetch before analysis.",
    )
    args = parser.parse_args()

    result = run_pipeline(
        workbook_path=Path(args.input),
        output_dir=Path(args.output_dir),
        cleaned_csv_path=Path(args.cleaned_csv),
        year=args.year,
        metadata_path=Path(args.metadata),
        fetch_live=args.live_fetch,
    )

    import pandas as pd
    from datetime import datetime, timedelta

    print(f"\n--- Pipeline Results ---")
    print(f"Cleaned rows: {result['cleaned_rows']}")
    print(f"Chart dates found: {result['date_count']}")
    print(f"Clean CSV: {result['cleaned_csv_path']}")
    print(f"Output directory: {result['output_dir']}")
    print(f"Report: {result['report_path']}")

    # Print Future Predictions Table
    future_pred_path = Path(result["output_dir"]) / "tables" / "top10_future_predictions.csv"
    if future_pred_path.exists():
        try:
            future_df = pd.read_csv(future_pred_path)
            if not future_df.empty:
                latest_date_str = future_df["Date"].iloc[0]
                latest_date = datetime.strptime(latest_date_str, "%Y-%m-%d").date()
                next_date = latest_date + timedelta(days=1)
                
                print(f"\n>>> PREDICTED TOP 10 FOR {next_date.strftime('%d %b %Y')} (Tomorrow)")
                print(f"{'Pred #':<7} | {'Song':<35} | {'Artist':<25} | {'Confidence'}")
                print("-" * 85)
                # Always show top 10 by today's rank; flag uncertain predictions with '?'
                top_10 = future_df.sort_values("Rank").head(10)
                for i, row in enumerate(top_10.itertuples(), 1):
                    song = str(row.Song)[:32] + "..." if len(str(row.Song)) > 35 else str(row.Song)
                    artist = str(row.Artist)[:22] + "..." if len(str(row.Artist)) > 25 else str(row.Artist)
                    prob = row.Predicted_Probability * 100
                    flag = "  (uncertain)" if prob < 50 else ""
                    print(f"{i:<7} | {song:<35} | {artist:<25} | {prob:>5.1f}%{flag}")
                print("-" * 85)
                print(f"Full table: {future_pred_path}")
        except Exception as e:
            print(f"\nNote: Could not print future predictions table: {e}")


if __name__ == "__main__":
    main()
