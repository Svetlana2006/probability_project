import argparse
from pathlib import Path

import pandas as pd

from spotify_pipeline import (
    analyze_artist_dominance,
    analyze_entries,
    analyze_genre,
    analyze_markov,
    analyze_prediction,
    analyze_rank_changes,
    analyze_survival,
    ensure_dirs,
    merge_genre_metadata,
    prepare_analysis_frame,
    save_hypothesis_table,
    save_rank_heatmap,
    write_report,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the Spotify chart analysis on a cleaned CSV.")
    parser.add_argument("--input", default="spotify_data.csv", help="Path to the cleaned CSV file.")
    parser.add_argument("--output-dir", default="output", help="Directory for analysis outputs.")
    parser.add_argument(
        "--metadata",
        default="genre_mapping.csv",
        help="Optional CSV file with Artist, Song, Genre columns.",
    )
    args = parser.parse_args()

    output_dirs = ensure_dirs(Path(args.output_dir))
    df = pd.read_csv(args.input)
    df_with_metadata, notes = merge_genre_metadata(df, Path(args.metadata), output_dirs["base"])
    analyzed, unique_dates = prepare_analysis_frame(df_with_metadata)
    analyzed.to_csv(output_dirs["tables"] / "spotify_analysis.csv", index=False)
    analyzed_public = analyzed.drop(columns=["Song_Key"])
    analyzed_public.to_csv(output_dirs["tables"] / "spotify_analysis_public.csv", index=False)
    analyzed_public.to_csv(output_dirs["base"] / "spotify_analysis.csv", index=False)

    if save_rank_heatmap(analyzed, output_dirs["figures"]) is None:
        notes.append("Rank heatmap was skipped because at least two chart dates are needed.")

    hypothesis_rows: list[dict] = []
    _, tests, more_notes = analyze_rank_changes(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    spells, tests, more_notes = analyze_survival(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    _, tests, more_notes = analyze_entries(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    _, tests, more_notes = analyze_artist_dominance(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    _, tests, more_notes = analyze_genre(analyzed, spells, output_dirs["tables"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    _, tests, more_notes = analyze_markov(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)
    _, tests, more_notes = analyze_prediction(analyzed, output_dirs["tables"], output_dirs["figures"])
    hypothesis_rows.extend(tests)
    notes.extend(more_notes)

    hypothesis_table = save_hypothesis_table(hypothesis_rows, output_dirs["tables"])
    write_report(
        report_path=output_dirs["base"] / "report.md",
        cleaned_rows=len(df),
        date_count=len(unique_dates),
        notes=notes,
        hypothesis_table=hypothesis_table,
        output_dirs=output_dirs,
    )

    print(f"Wrote analysis outputs to {output_dirs['base'].resolve()}")
    print(f"Report: {(output_dirs['base'] / 'report.md').resolve()}")


if __name__ == "__main__":
    main()
