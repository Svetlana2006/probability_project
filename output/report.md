# Spotify Top 50 Probability Pipeline Report

## Dataset Snapshot
- Cleaned chart rows: 50
- Unique chart dates: 1
- This pipeline is designed for daily Spotify Top 50 exports stored as one sheet per date.

## What Was Automated
- Cleaned the workbook into `spotify_data.csv`.
- Computed rank deltas, chart spells, artist concentration metrics, entry rates, Markov transitions, and Top 10 prediction features.
- Generated CSV tables in `output/tables` and figures in `output/figures`.
- Wrote a hypothesis-test summary table for the statistical questions that had enough data.

## Notes
- Genre metadata matched for 50 of 50 chart rows.
- Rank heatmap was skipped because at least two chart dates are needed.
- Rank-change analysis was skipped because the dataset has no consecutive chart observations yet.
- Survival distribution fitting was skipped because fewer than 2 completed survival spells were available.
- 100.00% of observed spells start on the first chart date in the dataset, so some song lifetimes may be left-truncated.
- Entry-rate analysis was skipped because at least two chart dates are needed.
- Pareto provides the lower AIC for artist appearance counts in this sample.
- Top 3 artists account for 28.00% of all Top 50 chart slots in the dataset.
- The artist-appearance Gini coefficient is 0.2674.
- A Cox proportional hazards model is not included because the required survival package is not installed in this environment.
- Markov analysis was skipped because the dataset has no next-day chart observations yet.
- Top 10 prediction was skipped because no day has a next-day label yet.

## Hypothesis Summary
- Artist Dominance: Chi-square goodness of fit -> Reject H0 (p=0.009239)
- Genre Momentum: Kruskal-Wallis -> Fail to reject H0 (p=NA)

## Output Files
- Tables:
  - artist_appearance_summary.csv
  - artist_dominance_metrics.csv
  - artist_frequency_distribution_fits.csv
  - genre_momentum_summary.csv
  - hypothesis_tests.csv
  - kaplan_meier_table.csv
  - new_entries_by_day.csv
  - spotify_analysis.csv
  - spotify_analysis_public.csv
  - survival_spells.csv
- Figures:
  - artist_lorenz_curve.png
  - survival_curve.png