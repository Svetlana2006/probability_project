import argparse
from pathlib import Path

from spotify_pipeline import clean_workbook, resolve_workbook_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Spotify chart workbook to a clean CSV.")
    parser.add_argument(
        "--input",
        default="Probability Project.xlsx",
        help="Path to the source Excel workbook.",
    )
    parser.add_argument(
        "--output",
        default="spotify_data.csv",
        help="Path to the cleaned CSV output.",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=2026,
        help="Year used when sheet names are in DDMM format.",
    )
    args = parser.parse_args()

    cleaned = clean_workbook(resolve_workbook_path(Path(args.input)), args.year)
    cleaned.to_csv(args.output, index=False)

    print(f"Cleaned {len(cleaned)} rows")
    print(f"Wrote {Path(args.output).resolve()}")
    print(cleaned.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
