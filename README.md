# Spotify Probability Pipeline

This project automates the workflow for your chart-survival study:

- append the latest India Top 50 daily chart into your workbook
- clean the Spotify Top 50 workbook
- build the analysis table
- run distribution, survival, entry-rate, artist-dominance, Markov, and prediction analyses
- generate plots, CSV tables, and a report

## 1. Append the latest daily chart

```powershell
python .\append_daily_chart.py
```

This command:

- opens the Spotify Charts India daily page in a persistent local browser session
- uses that browser session to fetch the India Top 50 chart data
- keeps the Top 50 rows
- writes them into `Probability Project.xlsx` as a sheet named `AUTO_YYYYMMDD`
- replaces that same auto sheet if you rerun it for the same chart date

If Spotify asks you to log in, complete it once in the opened browser window. The session is then reused from:

- `.\.browser-profile\spotify-charts`

For first-time setup, run it without `--headless`. After the session is saved, you can also use:

```powershell
python .\append_daily_chart.py --headless
```

## 2. Run the analysis

```powershell
python .\run_pipeline.py
```

Default input:

- workbook: `.\Probability Project.xlsx`
- cleaned CSV: `spotify_data.csv`
- outputs: `output\`
- optional genre metadata: `genre_mapping.csv`

## What the pipeline does

1. Cleaning

- Reads every sheet in the workbook as one chart date
- Detects the real header row even if the sheet has blank rows above it
- Also accepts auto-imported sheets that already contain `Date`, `Rank`, `Song`, and `Artist`
- Keeps only `Date`, `Rank`, `Song`, and `Artist`
- Removes bracket metadata like `(w/ ...)` and `(From "...")`
- Saves the cleaned data to `spotify_data.csv`

2. Core probability analysis

- Computes `Prev_Rank`, `Delta_R`, chart spells, entry flags, and state transitions
- Fits Normal and Laplace distributions to rank changes
- Fits Poisson and zero-inflated Poisson models to jump counts
- Builds Kaplan-Meier style survival tables and Geometric / Exponential / Weibull fits
- Counts new entries per day and runs Poisson-based tests
- Measures artist concentration with top-3 share, Gini, Lorenz curve, and distribution fits
- Builds a Markov transition matrix for `Top10`, `11-25`, `26-50`, and `Exit`
- Trains a logistic regression to predict whether a song stays in the Top 10 tomorrow

3. Outputs

- `output\report.md`
- `output\tables\*.csv`
- `output\figures\*.png`

## Genre analysis

The workbook does not contain genre labels, so the pipeline creates:

- `output\genre_mapping_template.csv`

To enable genre momentum analysis, copy that template to `genre_mapping.csv`, fill the `Genre` column, and rerun:

```powershell
python .\run_pipeline.py --metadata .\genre_mapping.csv
```

## Separate commands

If you want the steps separately:

```powershell
python .\clean_spotify_chart.py
python .\analyze_spotify_chart.py
```

## Source note

The default append command uses a browser-authenticated Spotify Charts session. A public Kworb fallback is still available if needed:

```powershell
python .\append_daily_chart.py --source kworb
```

## Current limitation

With only one chart date, the cleaning step works fully, but multi-day analyses such as rank changes, survival, transitions, and prediction will be skipped automatically until you add more daily sheets.
