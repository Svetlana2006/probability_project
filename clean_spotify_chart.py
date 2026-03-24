import argparse
import re
from pathlib import Path

import pandas as pd


def infer_date(sheet_name: str, year: int | None) -> str:
    text = str(sheet_name).strip()

    if re.fullmatch(r"\d{8}", text):
        return pd.to_datetime(text, format="%Y%m%d").strftime("%Y-%m-%d")

    if re.fullmatch(r"\d{4}", text):
        if year is None:
            raise ValueError(
                f"Sheet '{sheet_name}' looks like DDMM. Pass --year so the date can be inferred."
            )
        day = int(text[:2])
        month = int(text[2:])
        return pd.Timestamp(year=year, month=month, day=day).strftime("%Y-%m-%d")

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")

    raise ValueError(f"Could not infer a date from sheet '{sheet_name}'.")


def normalize_text(value: str) -> str:
    text = str(value)
    text = re.sub(r"\s*\([^)]*\)", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -")


def split_artist_title(value: str) -> tuple[str, str]:
    text = str(value).strip()
    if " - " in text:
        artist, song = text.split(" - ", 1)
    else:
        artist, song = "", text
    return normalize_text(artist), normalize_text(song)


def clean_sheet(workbook_path: Path, sheet_name: str, year: int | None) -> pd.DataFrame:
    raw = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
    header_row_index = None
    for idx, row in raw.iterrows():
        values = row.fillna("").astype(str).str.strip().tolist()
        if "Pos" in values and "Artist and Title" in values:
            header_row_index = idx
            break

    if header_row_index is None:
        raise ValueError(
            f"Sheet '{sheet_name}' does not contain the expected 'Pos' and 'Artist and Title' columns."
        )

    header = raw.iloc[header_row_index].fillna("").astype(str).str.strip()
    data = raw.iloc[header_row_index + 1 :].copy()
    data.columns = header

    cleaned = data[["Pos", "Artist and Title"]].copy()
    cleaned["Rank"] = pd.to_numeric(cleaned["Pos"], errors="coerce")
    cleaned = cleaned.dropna(subset=["Rank", "Artist and Title"])
    cleaned["Rank"] = cleaned["Rank"].astype(int)
    cleaned["Date"] = infer_date(sheet_name, year)

    artists, songs = zip(*cleaned["Artist and Title"].map(split_artist_title))
    cleaned["Artist"] = list(artists)
    cleaned["Song"] = list(songs)

    cleaned = cleaned[["Date", "Rank", "Song", "Artist"]]
    cleaned = cleaned.sort_values(["Date", "Rank"], kind="stable").reset_index(drop=True)
    return cleaned


def clean_workbook(workbook_path: Path, year: int | None) -> pd.DataFrame:
    excel_file = pd.ExcelFile(workbook_path)
    frames = [clean_sheet(workbook_path, sheet_name, year) for sheet_name in excel_file.sheet_names]
    combined = pd.concat(frames, ignore_index=True)
    return combined.sort_values(["Date", "Rank"], kind="stable").reset_index(drop=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Spotify chart workbook to a clean CSV.")
    parser.add_argument(
        "--input",
        default=r"c:\Users\Svetlana\Downloads\Probability Project.xlsx",
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

    workbook_path = Path(args.input)
    output_path = Path(args.output)

    cleaned = clean_workbook(workbook_path, args.year)
    cleaned.to_csv(output_path, index=False)

    print(f"Cleaned {len(cleaned)} rows from {workbook_path}")
    print(f"Wrote {output_path.resolve()}")
    print(cleaned.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
