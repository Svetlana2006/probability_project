import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from scipy import stats
from scipy.stats import shapiro

try:
    import seaborn as sns
except ModuleNotFoundError:
    sns = None


def get_state(rank: float) -> str:
    if pd.isna(rank):
        return "S4"
    if rank <= 10:
        return "S1"
    if rank <= 25:
        return "S2"
    return "S3"


def build_analysis(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values(by=["Song", "Date"], kind="stable").reset_index(drop=True)
    df["Prev_Rank"] = df.groupby("Song")["Rank"].shift(1)
    df["Delta_R"] = df["Rank"] - df["Prev_Rank"]
    df["State"] = df["Rank"].apply(get_state)
    df["Next_State"] = df.groupby("Song")["State"].shift(-1)
    return df


def save_transition_outputs(df: pd.DataFrame, output_dir: Path) -> None:
    transitions = df.dropna(subset=["Next_State"]).copy()
    if transitions.empty:
        print("Transition matrix skipped: you need at least two dates for the same song.")
        return

    transition_counts = pd.crosstab(transitions["State"], transitions["Next_State"])
    transition_probs = transition_counts.div(transition_counts.sum(axis=1), axis=0)

    counts_path = output_dir / "transition_counts.csv"
    probs_path = output_dir / "transition_matrix.csv"
    transition_counts.to_csv(counts_path)
    transition_probs.to_csv(probs_path)

    print(f"Saved transition counts to {counts_path.resolve()}")
    print(f"Saved transition probabilities to {probs_path.resolve()}")
    print(transition_probs.round(4).to_string())


def save_distribution_outputs(df: pd.DataFrame, output_dir: Path) -> None:
    delta = df["Delta_R"].dropna()
    if delta.empty:
        print("Distribution analysis skipped: no previous-day ranks are available yet.")
        return

    plot_path = output_dir / "delta_rank_distribution.png"
    plt.figure(figsize=(10, 6))
    if sns is not None:
        sns.histplot(delta, bins=30, kde=True)
    else:
        plt.hist(delta, bins=30, edgecolor="black")
    plt.title("Distribution of Rank Changes (Delta R)")
    plt.xlabel("Delta R")
    plt.tight_layout()
    plt.savefig(plot_path, dpi=150)
    plt.close()

    mu, sigma = stats.norm.fit(delta)
    loc, scale = stats.laplace.fit(delta)

    print(f"Saved distribution plot to {plot_path.resolve()}")
    print(f"Normal fit: mu={mu:.4f}, sigma={sigma:.4f}")
    print(f"Laplace fit: loc={loc:.4f}, scale={scale:.4f}")

    if len(delta) >= 3:
        stat, p_value = shapiro(delta)
        print(f"Shapiro-Wilk p-value: {p_value:.6f}")
        if p_value < 0.05:
            print("Interpretation: the delta distribution is not normal at the 5% level.")
        else:
            print("Interpretation: normality is not rejected at the 5% level.")
    else:
        print("Shapiro-Wilk test skipped: it needs at least 3 non-null delta values.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze cleaned Spotify chart data.")
    parser.add_argument(
        "--input",
        default="spotify_data.csv",
        help="Path to the cleaned CSV file.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Directory where analysis outputs will be written.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(input_path)
    analyzed = build_analysis(df)

    analyzed_path = output_dir / "spotify_analysis.csv"
    analyzed.to_csv(analyzed_path, index=False)
    print(f"Saved analysis table to {analyzed_path.resolve()}")
    print(analyzed.head(10).to_string(index=False))

    unique_dates = analyzed["Date"].nunique()
    print(f"Unique chart dates found: {unique_dates}")
    if unique_dates < 2:
        print("Only one chart date is available, so Delta_R and transitions will be mostly empty.")

    save_transition_outputs(analyzed, output_dir)
    save_distribution_outputs(analyzed, output_dir)


if __name__ == "__main__":
    main()
