import argparse
from pathlib import Path

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
    args = parser.parse_args()

    result = run_pipeline(
        workbook_path=Path(args.input),
        output_dir=Path(args.output_dir),
        cleaned_csv_path=Path(args.cleaned_csv),
        year=args.year,
        metadata_path=Path(args.metadata),
    )

    print(f"Cleaned rows: {result['cleaned_rows']}")
    print(f"Chart dates found: {result['date_count']}")
    print(f"Clean CSV: {result['cleaned_csv_path']}")
    print(f"Output directory: {result['output_dir']}")
    print(f"Report: {result['report_path']}")


if __name__ == "__main__":
    main()
