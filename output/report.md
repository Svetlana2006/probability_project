# Spotify Top 50 Probability Pipeline Report

## Dataset Snapshot
- Cleaned chart rows: 1050
- Unique chart dates: 21
- This pipeline is designed for daily Spotify Top 50 exports stored as one sheet per date.

## What Was Automated
- Cleaned the workbook into `spotify_data.csv`.
- Computed rank deltas, chart spells, artist concentration metrics, entry rates, Markov transitions, and Top 10 prediction features.
- Generated CSV tables in `output/tables` and figures in `output/figures`.
- Wrote a hypothesis-test summary table for the statistical questions that had enough data.

## Notes
- Genre metadata matched for 1050 of 1050 chart rows.
- Laplace provides a better AIC than Normal, which supports heavier-tailed rank movement.
- Observed zero-move share in |Delta_R| is 28.89%; the zero-inflated Poisson fit checks whether no-move days exceed Poisson expectations.
- Survival fits were estimated on completed spells only; spells still active at the final chart date are treated as right-censored in the Kaplan-Meier table.
- 47.62% of observed spells start on the first chart date in the dataset, so some song lifetimes may be left-truncated.
- Lognormal provides the lower AIC for artist appearance counts in this sample.
- Top 3 artists account for 30.00% of all Top 50 chart slots in the dataset.
- The artist-appearance Gini coefficient is 0.4392.
- A Cox proportional hazards model is not included because the required survival package is not installed in this environment.
- Expected time before exit was computed from the absorbing Markov chain fundamental matrix.
- The full absorbing chain has stationary mass on Exit; a conditional stationary distribution among surviving states was also saved.
- Generated future Top 10 predictions for the latest data (2026-04-04). Rank_Velocity (3-day rolling avg rank change, negative = climbing) is now a model feature. 0 song(s) currently outside the Top 10 are predicted to enter it tomorrow (Is_Top10_Challenger=1 in top10_future_predictions.csv).
- Top 10 prediction uses a chronological train/test split to measure next-day forecasting. Rank_Velocity (3-day rolling mean of daily rank change; negative = climbing) is included as a feature so fast-rising songs outside the Top 10 are scored as potential challengers.

## Hypothesis Summary
- Rank Changes: Shapiro-Wilk -> Reject H0 (p=0.000000)
- Rank Changes: Kolmogorov-Smirnov -> Reject H0 (p=0.000000)
- Rank Changes: Kolmogorov-Smirnov -> Reject H0 (p=0.000000)
- Survival: Likelihood Ratio Test (Exponential vs Weibull) -> Fail to reject H0 (p=0.222568)
- New Entries: Chi-square goodness of fit -> Fail to reject H0 (p=0.770876)
- New Entries: Two-sample Poisson rate z-test -> Fail to reject H0 (p=0.596599)
- Artist Dominance: Chi-square goodness of fit -> Reject H0 (p=0.000000)
- Genre Momentum: Kruskal-Wallis -> Reject H0 (p=0.007830)
- Top10 Prediction: Logistic regression coefficient sign -> Interpret coefficient directly (p=NA)
- Top10 Prediction: Logistic regression coefficient sign -> Interpret coefficient directly (p=NA)

## Output Files
- Tables:
  - artist_appearance_summary.csv
  - artist_dominance_metrics.csv
  - artist_frequency_distribution_fits.csv
  - genre_momentum_summary.csv
  - hypothesis_tests.csv
  - kaplan_meier_table.csv
  - markov_absorption_summary.csv
  - markov_conditional_stationary_distribution.csv
  - markov_transition_counts.csv
  - markov_transition_matrix.csv
  - new_entries_by_day.csv
  - rank_change_distribution_fits.csv
  - spotify_analysis.csv
  - spotify_analysis_public.csv
  - survival_distribution_fits.csv
  - survival_spells.csv
  - top10_future_predictions.csv
  - top10_prediction_coefficients.csv
  - top10_prediction_metrics.csv
  - top10_prediction_scored_rows.csv
- Figures:
  - artist_lorenz_curve.png
  - delta_rank_distribution.png
  - delta_rank_qqplot.png
  - markov_transition_matrix.png
  - new_entries_by_day.png
  - rank_heatmap.png
  - survival_curve.png
  - top10_prediction_coefficients.png
  - top10_prediction_confusion_matrix.png