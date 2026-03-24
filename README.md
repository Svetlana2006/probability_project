# Spotify Probability Pipeline

This project turns the raw chart workbook into a clean CSV, then computes the first analysis fields you need for rank-change and Markov-state work.

## 1. Clean the workbook

```powershell
python .\clean_spotify_chart.py
```

What this does:

- Reads the Excel workbook from `C:\Users\Svetlana\Downloads\Probability Project.xlsx`
- Uses the first row inside each sheet as the real header row
- Keeps only `Date`, `Rank`, `Song`, and `Artist`
- Ignores the position-change column like `=`, `+1`, `-2`
- Removes bracket metadata such as `(w/ ...)` and `(From "...")`
- Saves the result to `spotify_data.csv`

## 2. Run the analysis

```powershell
python .\analyze_spotify_chart.py
```

What this does:

- Sorts each song by date
- Computes `Prev_Rank`
- Computes `Delta_R = Rank - Prev_Rank`
- Maps ranks into states:
  - `S1`: ranks 1 to 10
  - `S2`: ranks 11 to 25
  - `S3`: ranks 26+
  - `S4`: missing rank
- Builds transition counts and transition probabilities when enough dates exist
- Saves the full table to `output\spotify_analysis.csv`
- Saves the histogram to `output\delta_rank_distribution.png`

## 3. Current limitation

The workbook you provided currently has one sheet for `2026-03-22`, so the cleaning step works now, but the analysis step does not yet have enough history for meaningful `Delta_R` or transition probabilities.

Once you add more days, rerun the same two commands.
