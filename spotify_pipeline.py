from __future__ import annotations

import math
import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import optimize, stats
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
)


STATE_ORDER = ["Top10", "11-25", "26-50", "Exit"]


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
    return cleaned.sort_values(["Date", "Rank"], kind="stable").reset_index(drop=True)


def clean_workbook(workbook_path: Path, year: int | None) -> pd.DataFrame:
    excel_file = pd.ExcelFile(workbook_path)
    frames = [clean_sheet(workbook_path, sheet_name, year) for sheet_name in excel_file.sheet_names]
    combined = pd.concat(frames, ignore_index=True)
    return combined.sort_values(["Date", "Rank"], kind="stable").reset_index(drop=True)


def add_song_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Song_Key"] = (
        df["Artist"].fillna("").astype(str).str.strip().str.lower()
        + " || "
        + df["Song"].fillna("").astype(str).str.strip().str.lower()
    )
    return df


def resolve_workbook_path(workbook_path: Path) -> Path:
    if workbook_path.exists():
        return workbook_path

    candidates = [
        Path.cwd() / workbook_path.name,
        Path.cwd() / "Probability Project.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise FileNotFoundError(f"Workbook not found: {workbook_path}")


def assign_state(rank: float) -> str:
    if pd.isna(rank):
        return "Exit"
    if rank <= 10:
        return "Top10"
    if rank <= 25:
        return "11-25"
    return "26-50"


def ensure_dirs(base_output_dir: Path) -> dict[str, Path]:
    tables_dir = base_output_dir / "tables"
    figures_dir = base_output_dir / "figures"
    base_output_dir.mkdir(parents=True, exist_ok=True)
    tables_dir.mkdir(parents=True, exist_ok=True)
    figures_dir.mkdir(parents=True, exist_ok=True)
    return {"base": base_output_dir, "tables": tables_dir, "figures": figures_dir}


def merge_genre_metadata(df: pd.DataFrame, metadata_path: Path, output_dir: Path) -> tuple[pd.DataFrame, list[str]]:
    df = add_song_keys(df)
    notes: list[str] = []
    template = (
        df[["Artist", "Song"]]
        .drop_duplicates()
        .sort_values(["Artist", "Song"], kind="stable")
        .assign(Genre="")
    )
    template_path = output_dir / "genre_mapping_template.csv"
    template.to_csv(template_path, index=False)

    if not metadata_path.exists():
        df["Genre"] = pd.NA
        notes.append(
            "Genre analysis was skipped because no genre mapping file was found. "
            f"Fill {template_path.name} and rerun to enable genre tests."
        )
        return df, notes

    metadata = pd.read_csv(metadata_path)
    required = {"Artist", "Song", "Genre"}
    if not required.issubset(metadata.columns):
        df["Genre"] = pd.NA
        notes.append(
            f"Genre analysis was skipped because {metadata_path.name} must contain Artist, Song, and Genre columns."
        )
        return df, notes

    metadata = add_song_keys(metadata[["Artist", "Song", "Genre"]].copy())
    metadata["Genre"] = metadata["Genre"].replace("", pd.NA)
    merged = df.merge(metadata[["Song_Key", "Genre"]], on="Song_Key", how="left")
    matched = merged["Genre"].notna().sum()
    notes.append(f"Genre metadata matched for {matched} of {len(merged)} chart rows.")
    return merged, notes


def compute_group_expanding_std(values: pd.Series) -> pd.Series:
    result: list[float] = []
    history: list[float] = []
    for value in values:
        if pd.notna(value):
            history.append(float(value))
        if len(history) < 2:
            result.append(0.0)
        else:
            result.append(float(np.std(history, ddof=0)))
    return pd.Series(result, index=values.index)


def prepare_analysis_frame(df: pd.DataFrame) -> tuple[pd.DataFrame, list[pd.Timestamp]]:
    df = add_song_keys(df)
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values(["Date", "Rank", "Song_Key"], kind="stable").reset_index(drop=True)

    unique_dates = sorted(df["Date"].drop_duplicates().tolist())
    date_to_index = {date: idx for idx, date in enumerate(unique_dates)}
    next_date_map = {unique_dates[idx]: unique_dates[idx + 1] for idx in range(len(unique_dates) - 1)}

    df["Date_Index"] = df["Date"].map(date_to_index)
    df["State"] = df["Rank"].apply(assign_state)
    df = df.sort_values(["Song_Key", "Date_Index"], kind="stable").reset_index(drop=True)

    prev_date_index = df.groupby("Song_Key")["Date_Index"].shift(1)
    prev_rank_raw = df.groupby("Song_Key")["Rank"].shift(1)
    df["New_Spell"] = prev_date_index.isna() | ((df["Date_Index"] - prev_date_index) != 1)
    df["Spell_Number"] = df.groupby("Song_Key")["New_Spell"].cumsum().astype(int)
    df["Days_On_Chart"] = df.groupby(["Song_Key", "Spell_Number"]).cumcount() + 1
    df["Total_Appearances"] = df.groupby("Song_Key").cumcount() + 1
    df["Artist_Past_Appearances"] = df.groupby("Artist").cumcount()
    df["Prev_Rank"] = np.where(df["New_Spell"], np.nan, prev_rank_raw)
    df["Delta_R"] = df["Rank"] - df["Prev_Rank"]
    df["Volatility_So_Far"] = (
        df.groupby(["Song_Key", "Spell_Number"], group_keys=False)["Delta_R"].apply(compute_group_expanding_std)
    )
    df["Is_New_Entry"] = (df["Days_On_Chart"] == 1) & (df["Date_Index"] > 0)

    if next_date_map:
        date_map_df = pd.DataFrame(
            {"Date": list(next_date_map.keys()), "Next_Date": list(next_date_map.values())}
        )
        date_map_df["Date"] = pd.to_datetime(date_map_df["Date"])
        date_map_df["Next_Date"] = pd.to_datetime(date_map_df["Next_Date"])
        df = df.merge(date_map_df, on="Date", how="left")
    else:
        df["Next_Date"] = pd.NaT
    df["Next_Date"] = pd.to_datetime(df["Next_Date"])

    next_lookup = df[["Song_Key", "Date", "Rank", "State"]].rename(
        columns={"Date": "Next_Date", "Rank": "Next_Rank", "State": "Observed_Next_State"}
    )
    next_lookup["Next_Date"] = pd.to_datetime(next_lookup["Next_Date"])
    df = df.merge(next_lookup, on=["Song_Key", "Next_Date"], how="left")
    has_follow_up = df["Next_Date"].notna()
    df["Next_State"] = np.where(has_follow_up, df["Observed_Next_State"].fillna("Exit"), pd.NA)
    df["Top10_Tomorrow"] = np.where(has_follow_up, (df["Next_Rank"] <= 10).fillna(False).astype(int), pd.NA)
    df["Exit_Tomorrow"] = np.where(has_follow_up, df["Observed_Next_State"].isna().astype(int), pd.NA)

    return df.sort_values(["Date", "Rank"], kind="stable").reset_index(drop=True), unique_dates


def save_rank_heatmap(df: pd.DataFrame, figures_dir: Path) -> str | None:
    unique_dates = df["Date"].nunique()
    if unique_dates < 2:
        return None

    top_songs = (
        df.groupby(["Song_Key", "Artist", "Song"])
        .agg(Appearances=("Date", "nunique"), Best_Rank=("Rank", "min"))
        .reset_index()
        .sort_values(["Appearances", "Best_Rank"], ascending=[False, True], kind="stable")
        .head(20)
    )
    top_keys = top_songs["Song_Key"].tolist()
    labels = [f"{row.Artist} - {row.Song}" for row in top_songs.itertuples()]
    pivot = (
        df[df["Song_Key"].isin(top_keys)]
        .pivot_table(index="Song_Key", columns="Date", values="Rank", aggfunc="first")
        .reindex(top_keys)
    )

    fig, ax = plt.subplots(figsize=(max(8, unique_dates * 0.7), max(6, len(top_keys) * 0.35)))
    matrix = np.ma.masked_invalid(pivot.to_numpy(dtype=float))
    cmap = plt.cm.YlOrRd_r.copy()
    cmap.set_bad(color="#f0f0f0")
    image = ax.imshow(matrix, aspect="auto", cmap=cmap, vmin=1, vmax=50)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=8)
    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels([pd.Timestamp(col).strftime("%m-%d") for col in pivot.columns], rotation=45, ha="right")
    ax.set_title("Rank Heatmap for Songs with the Most Appearances")
    fig.colorbar(image, ax=ax, label="Rank")
    fig.tight_layout()
    output_path = figures_dir / "rank_heatmap.png"
    fig.savefig(output_path, dpi=150)
    plt.close(fig)
    return output_path.name


def fit_zero_inflated_poisson(values: pd.Series) -> tuple[float, float, float]:
    x = values.astype(int).to_numpy()
    mean_x = float(np.mean(x)) if len(x) else 0.1
    zero_share = float(np.mean(x == 0)) if len(x) else 0.0
    init_pi = min(max(zero_share - math.exp(-max(mean_x, 1e-6)), 0.01), 0.95)
    init_lambda = max(mean_x, 1e-3)

    def neg_log_likelihood(params: np.ndarray) -> float:
        log_lambda, logit_pi = params
        lam = float(np.exp(log_lambda))
        pi = 1 / (1 + math.exp(-float(logit_pi)))
        zero_prob = pi + (1 - pi) * np.exp(-lam)
        positive_prob = (1 - pi) * np.exp(-lam) * np.power(lam, x) / stats.factorial(x)
        probs = np.where(x == 0, zero_prob, positive_prob)
        probs = np.clip(probs, 1e-12, None)
        return -float(np.sum(np.log(probs)))

    result = optimize.minimize(
        neg_log_likelihood,
        x0=np.array([math.log(init_lambda), math.log(init_pi / (1 - init_pi))]),
        method="L-BFGS-B",
    )
    log_lambda, logit_pi = result.x
    lam = float(np.exp(log_lambda))
    pi = 1 / (1 + math.exp(-float(logit_pi)))
    return lam, float(pi), -float(result.fun)


def analyze_rank_changes(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    delta = df["Delta_R"].dropna().astype(float)
    jump_count = delta.abs().astype(int)

    if delta.empty:
        notes.append("Rank-change analysis was skipped because the dataset has no consecutive chart observations yet.")
        return pd.DataFrame(), tests, notes

    mu, sigma = stats.norm.fit(delta)
    laplace_loc, laplace_scale = stats.laplace.fit(delta)
    sigma = max(sigma, 1e-9)
    laplace_scale = max(laplace_scale, 1e-9)
    normal_ll = float(np.sum(stats.norm.logpdf(delta, mu, sigma)))
    laplace_ll = float(np.sum(stats.laplace.logpdf(delta, laplace_loc, laplace_scale)))
    poisson_lambda = max(float(jump_count.mean()), 1e-9)
    poisson_ll = float(np.sum(stats.poisson.logpmf(jump_count, poisson_lambda)))
    zip_lambda, zip_pi, zip_ll = fit_zero_inflated_poisson(jump_count)

    fit_df = pd.DataFrame(
        [
            {
                "Model": "Normal",
                "Target": "Delta_R",
                "LogLikelihood": normal_ll,
                "AIC": 2 * 2 - 2 * normal_ll,
                "Parameters": f"mu={mu:.4f}, sigma={sigma:.4f}",
            },
            {
                "Model": "Laplace",
                "Target": "Delta_R",
                "LogLikelihood": laplace_ll,
                "AIC": 2 * 2 - 2 * laplace_ll,
                "Parameters": f"loc={laplace_loc:.4f}, scale={laplace_scale:.4f}",
            },
            {
                "Model": "Poisson",
                "Target": "|Delta_R|",
                "LogLikelihood": poisson_ll,
                "AIC": 2 * 1 - 2 * poisson_ll,
                "Parameters": f"lambda={poisson_lambda:.4f}",
            },
            {
                "Model": "ZeroInflatedPoisson",
                "Target": "|Delta_R|",
                "LogLikelihood": zip_ll,
                "AIC": 2 * 2 - 2 * zip_ll,
                "Parameters": f"lambda={zip_lambda:.4f}, zero_inflation={zip_pi:.4f}",
            },
        ]
    )
    fit_df.to_csv(tables_dir / "rank_change_distribution_fits.csv", index=False)

    fig, ax = plt.subplots(figsize=(9, 6))
    bins = min(30, max(10, int(np.sqrt(len(delta)))))
    ax.hist(delta, bins=bins, density=True, alpha=0.6, color="#4c78a8", edgecolor="white")
    xs = np.linspace(delta.min() - 1, delta.max() + 1, 400)
    ax.plot(xs, stats.norm.pdf(xs, mu, sigma), label="Normal fit", color="#f58518", linewidth=2)
    ax.plot(xs, stats.laplace.pdf(xs, laplace_loc, laplace_scale), label="Laplace fit", color="#54a24b", linewidth=2)
    ax.set_title("Distribution of Daily Rank Changes")
    ax.set_xlabel("Delta_R")
    ax.set_ylabel("Density")
    ax.legend()
    fig.tight_layout()
    fig.savefig(figures_dir / "delta_rank_distribution.png", dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 7))
    stats.probplot(delta, dist="norm", plot=ax)
    ax.set_title("QQ Plot for Rank Changes vs Normal")
    fig.tight_layout()
    fig.savefig(figures_dir / "delta_rank_qqplot.png", dpi=150)
    plt.close(fig)

    if len(delta) >= 3:
        shapiro_stat, shapiro_p = stats.shapiro(delta)
        tests.append(
            {
                "Analysis": "Rank Changes",
                "Hypothesis": "H0: Delta_R follows a normal distribution",
                "Test": "Shapiro-Wilk",
                "Statistic": shapiro_stat,
                "PValue": shapiro_p,
                "Decision_5pct": "Reject H0" if shapiro_p < 0.05 else "Fail to reject H0",
                "Note": "Applied to non-null day-to-day rank changes.",
            }
        )
    else:
        notes.append("Shapiro-Wilk test was skipped because fewer than 3 rank changes were available.")

    ks_normal_stat, ks_normal_p = stats.kstest(delta, "norm", args=(mu, sigma))
    tests.append(
        {
            "Analysis": "Rank Changes",
            "Hypothesis": "H0: Delta_R follows the fitted normal distribution",
            "Test": "Kolmogorov-Smirnov",
            "Statistic": ks_normal_stat,
            "PValue": ks_normal_p,
            "Decision_5pct": "Reject H0" if ks_normal_p < 0.05 else "Fail to reject H0",
            "Note": "Compared against the normal fit estimated from the sample.",
        }
    )

    ks_laplace_stat, ks_laplace_p = stats.kstest(delta, "laplace", args=(laplace_loc, laplace_scale))
    tests.append(
        {
            "Analysis": "Rank Changes",
            "Hypothesis": "H0: Delta_R follows the fitted Laplace distribution",
            "Test": "Kolmogorov-Smirnov",
            "Statistic": ks_laplace_stat,
            "PValue": ks_laplace_p,
            "Decision_5pct": "Reject H0" if ks_laplace_p < 0.05 else "Fail to reject H0",
            "Note": "Compared against the Laplace fit estimated from the sample.",
        }
    )

    normal_aic = fit_df.loc[fit_df["Model"] == "Normal", "AIC"].iloc[0]
    laplace_aic = fit_df.loc[fit_df["Model"] == "Laplace", "AIC"].iloc[0]
    if laplace_aic < normal_aic:
        notes.append("Laplace provides a better AIC than Normal, which supports heavier-tailed rank movement.")
    else:
        notes.append("Normal provides an equal or better AIC than Laplace on the observed rank changes.")
    zero_share = float(np.mean(jump_count == 0))
    notes.append(
        f"Observed zero-move share in |Delta_R| is {zero_share:.2%}; the zero-inflated Poisson fit checks whether no-move days exceed Poisson expectations."
    )
    return fit_df, tests, notes


def build_survival_spells(df: pd.DataFrame) -> pd.DataFrame:
    spells = (
        df.groupby(["Song_Key", "Spell_Number"])
        .agg(
            Artist=("Artist", "first"),
            Song=("Song", "first"),
            Genre=("Genre", "first"),
            Start_Date=("Date", "min"),
            End_Date=("Date", "max"),
            Duration=("Date", "size"),
            Start_Index=("Date_Index", "min"),
            End_Index=("Date_Index", "max"),
        )
        .reset_index()
    )
    max_index = int(df["Date_Index"].max())
    spells["Event_Observed"] = (spells["End_Index"] < max_index).astype(int)
    spells["Left_Boundary_Spell"] = (spells["Start_Index"] == 0).astype(int)
    return spells


def kaplan_meier_table(durations: pd.Series, events: pd.Series) -> pd.DataFrame:
    km_rows: list[dict] = []
    survival = 1.0
    at_risk = len(durations)
    for time in sorted(durations.unique()):
        d_i = int(((durations == time) & (events == 1)).sum())
        c_i = int(((durations == time) & (events == 0)).sum())
        if at_risk > 0:
            survival *= 1 - (d_i / at_risk)
        km_rows.append(
            {
                "Duration": int(time),
                "At_Risk": int(at_risk),
                "Events": d_i,
                "Censored": c_i,
                "Survival_Probability": survival,
            }
        )
        at_risk -= d_i + c_i
    return pd.DataFrame(km_rows)


def analyze_survival(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    spells = build_survival_spells(df)
    spells.to_csv(tables_dir / "survival_spells.csv", index=False)

    if spells.empty:
        notes.append("Survival analysis was skipped because no spells could be constructed.")
        return spells, tests, notes

    km = kaplan_meier_table(spells["Duration"], spells["Event_Observed"])
    km.to_csv(tables_dir / "kaplan_meier_table.csv", index=False)

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.step([0] + km["Duration"].tolist(), [1.0] + km["Survival_Probability"].tolist(), where="post", color="#4c78a8")
    ax.set_title("Kaplan-Meier Style Survival Curve for Top 50 Presence")
    ax.set_xlabel("Consecutive days in chart")
    ax.set_ylabel("Survival probability")
    ax.set_ylim(0, 1.05)
    fig.tight_layout()
    fig.savefig(figures_dir / "survival_curve.png", dpi=150)
    plt.close(fig)

    complete = spells.loc[spells["Event_Observed"] == 1, "Duration"].astype(float)
    if len(complete) >= 2:
        geom_p = min(max(1.0 / complete.mean(), 1e-9), 1 - 1e-9)
        geom_ll = float(np.sum(stats.geom.logpmf(complete.astype(int), geom_p)))
        exp_scale = max(float(complete.mean()), 1e-9)
        exp_ll = float(np.sum(stats.expon.logpdf(complete, scale=exp_scale)))
        weib_shape, _, weib_scale = stats.weibull_min.fit(complete, floc=0)
        weib_shape = max(float(weib_shape), 1e-9)
        weib_scale = max(float(weib_scale), 1e-9)
        weib_ll = float(np.sum(stats.weibull_min.logpdf(complete, weib_shape, loc=0, scale=weib_scale)))
        fit_df = pd.DataFrame(
            [
                {
                    "Model": "Geometric",
                    "LogLikelihood": geom_ll,
                    "AIC": 2 * 1 - 2 * geom_ll,
                    "Parameters": f"p={geom_p:.4f}",
                },
                {
                    "Model": "Exponential",
                    "LogLikelihood": exp_ll,
                    "AIC": 2 * 1 - 2 * exp_ll,
                    "Parameters": f"scale={exp_scale:.4f}",
                },
                {
                    "Model": "Weibull",
                    "LogLikelihood": weib_ll,
                    "AIC": 2 * 2 - 2 * weib_ll,
                    "Parameters": f"shape={weib_shape:.4f}, scale={weib_scale:.4f}",
                },
            ]
        )
        fit_df.to_csv(tables_dir / "survival_distribution_fits.csv", index=False)

        lr_stat = 2 * (weib_ll - exp_ll)
        lr_p = 1 - stats.chi2.cdf(lr_stat, df=1)
        tests.append(
            {
                "Analysis": "Survival",
                "Hypothesis": "H0: Song survival is memoryless",
                "Test": "Likelihood Ratio Test (Exponential vs Weibull)",
                "Statistic": lr_stat,
                "PValue": lr_p,
                "Decision_5pct": "Reject H0" if lr_p < 0.05 else "Fail to reject H0",
                "Note": "Weibull shape > 1 suggests persistence strengthens with time; < 1 suggests early drop-off risk.",
            }
        )
        notes.append(
            "Survival fits were estimated on completed spells only; spells still active at the final chart date are treated as right-censored in the Kaplan-Meier table."
        )
    else:
        notes.append("Survival distribution fitting was skipped because fewer than 2 completed survival spells were available.")

    left_boundary_share = float(spells["Left_Boundary_Spell"].mean())
    notes.append(
        f"{left_boundary_share:.2%} of observed spells start on the first chart date in the dataset, so some song lifetimes may be left-truncated."
    )
    return spells, tests, notes


def poisson_rate_test(counts_a: pd.Series, counts_b: pd.Series) -> tuple[float, float]:
    total_a = float(counts_a.sum())
    total_b = float(counts_b.sum())
    exposure_a = max(float(len(counts_a)), 1.0)
    exposure_b = max(float(len(counts_b)), 1.0)
    rate_a = total_a / exposure_a
    rate_b = total_b / exposure_b
    if total_a == 0 and total_b == 0:
        return 0.0, 1.0
    se = math.sqrt((total_a / (exposure_a ** 2 + 1e-12)) + (total_b / (exposure_b ** 2 + 1e-12)))
    if se == 0:
        return 0.0, 1.0
    z_stat = (rate_a - rate_b) / se
    p_value = 2 * (1 - stats.norm.cdf(abs(z_stat)))
    return z_stat, p_value


def analyze_entries(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    daily = (
        df.groupby("Date")
        .agg(
            New_Entries=("Is_New_Entry", "sum"),
            Chart_Size=("Song_Key", "size"),
        )
        .reset_index()
    )
    daily["Weekday"] = daily["Date"].dt.day_name()
    daily["Is_Weekend"] = daily["Date"].dt.dayofweek >= 5
    daily.to_csv(tables_dir / "new_entries_by_day.csv", index=False)

    if len(daily) <= 1:
        notes.append("Entry-rate analysis was skipped because at least two chart dates are needed.")
        return daily, tests, notes

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(daily["Date"].dt.strftime("%Y-%m-%d"), daily["New_Entries"], color="#72b7b2")
    ax.set_title("New Songs Entering the Top 50 by Day")
    ax.set_xlabel("Date")
    ax.set_ylabel("New entries")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    fig.savefig(figures_dir / "new_entries_by_day.png", dpi=150)
    plt.close(fig)

    counts = daily["New_Entries"].astype(int)
    lam = float(counts.mean())
    value_counts = counts.value_counts().sort_index()
    observed = value_counts.to_numpy(dtype=float)
    support = value_counts.index.to_numpy(dtype=int)
    expected = stats.poisson.pmf(support, lam) * len(counts)
    if expected.sum() < len(counts):
        expected[-1] += len(counts) - expected.sum()

    if len(observed) >= 2 and np.all(expected > 0):
        chi_stat = float(np.sum(((observed - expected) ** 2) / expected))
        dof = max(len(observed) - 2, 1)
        chi_p = 1 - stats.chi2.cdf(chi_stat, dof)
        tests.append(
            {
                "Analysis": "New Entries",
                "Hypothesis": "H0: Daily new-entry counts follow a Poisson process with constant rate",
                "Test": "Chi-square goodness of fit",
                "Statistic": chi_stat,
                "PValue": chi_p,
                "Decision_5pct": "Reject H0" if chi_p < 0.05 else "Fail to reject H0",
                "Note": "Poisson rate estimated from the observed daily entry counts.",
            }
        )
    else:
        notes.append("Poisson goodness-of-fit test was skipped because the daily entry-count support was too sparse.")

    weekend = daily.loc[daily["Is_Weekend"], "New_Entries"]
    weekday = daily.loc[~daily["Is_Weekend"], "New_Entries"]
    if len(weekend) >= 1 and len(weekday) >= 1:
        z_stat, p_value = poisson_rate_test(weekend, weekday)
        tests.append(
            {
                "Analysis": "New Entries",
                "Hypothesis": "H0: Weekend and weekday entry rates are equal",
                "Test": "Two-sample Poisson rate z-test",
                "Statistic": z_stat,
                "PValue": p_value,
                "Decision_5pct": "Reject H0" if p_value < 0.05 else "Fail to reject H0",
                "Note": "Compares the mean number of new entrants per weekend day against weekday days.",
            }
        )
    else:
        notes.append("Weekend versus weekday entry-rate comparison was skipped because one of the groups had no observations.")

    return daily, tests, notes


def gini_coefficient(values: np.ndarray) -> float:
    if len(values) == 0:
        return float("nan")
    sorted_values = np.sort(values.astype(float))
    n = len(sorted_values)
    total = sorted_values.sum()
    if total == 0:
        return 0.0
    index = np.arange(1, n + 1)
    return float((2 * np.sum(index * sorted_values) / (n * total)) - (n + 1) / n)


def analyze_artist_dominance(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    artist_counts = (
        df.groupby("Artist")
        .agg(Appearances=("Song_Key", "size"), Unique_Songs=("Song_Key", "nunique"), Best_Rank=("Rank", "min"))
        .reset_index()
        .sort_values(["Appearances", "Best_Rank"], ascending=[False, True], kind="stable")
    )
    artist_counts.to_csv(tables_dir / "artist_appearance_summary.csv", index=False)

    if artist_counts.empty:
        notes.append("Artist-dominance analysis was skipped because no artist rows were available.")
        return artist_counts, tests, notes

    values = artist_counts["Appearances"].to_numpy(dtype=float)
    top3_share = float(values[:3].sum() / values.sum())
    gini = gini_coefficient(values)
    summary = pd.DataFrame(
        [{"Metric": "Top3_Share", "Value": top3_share}, {"Metric": "Gini", "Value": gini}]
    )
    summary.to_csv(tables_dir / "artist_dominance_metrics.csv", index=False)

    cumulative_artists = np.linspace(0, 1, len(values) + 1)
    cumulative_share = np.concatenate([[0.0], np.cumsum(np.sort(values)) / np.sum(values)])
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.plot(cumulative_artists, cumulative_share, label="Lorenz curve", color="#4c78a8", linewidth=2)
    ax.plot([0, 1], [0, 1], linestyle="--", color="#999999", label="Equality line")
    ax.set_title("Artist Dominance Lorenz Curve")
    ax.set_xlabel("Cumulative share of artists")
    ax.set_ylabel("Cumulative share of appearances")
    ax.legend()
    fig.tight_layout()
    fig.savefig(figures_dir / "artist_lorenz_curve.png", dpi=150)
    plt.close(fig)

    expected = np.repeat(values.sum() / len(values), len(values))
    chi_stat = float(np.sum(((values - expected) ** 2) / expected))
    chi_p = 1 - stats.chi2.cdf(chi_stat, max(len(values) - 1, 1))
    tests.append(
        {
            "Analysis": "Artist Dominance",
            "Hypothesis": "H0: Artist chart appearances are uniformly distributed",
            "Test": "Chi-square goodness of fit",
            "Statistic": chi_stat,
            "PValue": chi_p,
            "Decision_5pct": "Reject H0" if chi_p < 0.05 else "Fail to reject H0",
            "Note": "Expected counts assume equal chart exposure across artists.",
        }
    )

    if len(values) >= 3:
        log_values = values[values > 0]
        lognorm_shape, lognorm_loc, lognorm_scale = stats.lognorm.fit(log_values, floc=0)
        pareto_shape, pareto_loc, pareto_scale = stats.pareto.fit(log_values, floc=0)
        lognorm_ll = float(np.sum(stats.lognorm.logpdf(log_values, lognorm_shape, loc=lognorm_loc, scale=lognorm_scale)))
        pareto_ll = float(np.sum(stats.pareto.logpdf(log_values, pareto_shape, loc=pareto_loc, scale=pareto_scale)))
        fit_df = pd.DataFrame(
            [
                {
                    "Model": "Lognormal",
                    "LogLikelihood": lognorm_ll,
                    "AIC": 2 * 2 - 2 * lognorm_ll,
                    "Parameters": f"shape={lognorm_shape:.4f}, scale={lognorm_scale:.4f}",
                },
                {
                    "Model": "Pareto",
                    "LogLikelihood": pareto_ll,
                    "AIC": 2 * 2 - 2 * pareto_ll,
                    "Parameters": f"shape={pareto_shape:.4f}, scale={pareto_scale:.4f}",
                },
            ]
        )
        fit_df.to_csv(tables_dir / "artist_frequency_distribution_fits.csv", index=False)
        preferred = fit_df.sort_values("AIC").iloc[0]["Model"]
        notes.append(f"{preferred} provides the lower AIC for artist appearance counts in this sample.")

    notes.append(f"Top 3 artists account for {top3_share:.2%} of all Top 50 chart slots in the dataset.")
    notes.append(f"The artist-appearance Gini coefficient is {gini:.4f}.")
    return artist_counts, tests, notes


def analyze_genre(df: pd.DataFrame, spells: pd.DataFrame, tables_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []

    if "Genre" not in df.columns or df["Genre"].dropna().empty:
        notes.append("Genre momentum analysis was skipped because genre labels are not available.")
        return pd.DataFrame(), tests, notes

    genre_spells = spells.dropna(subset=["Genre"]).copy()
    if genre_spells["Genre"].nunique() < 2:
        notes.append("Genre momentum analysis was skipped because fewer than two genre groups were available.")
        return pd.DataFrame(), tests, notes

    summary = (
        genre_spells.groupby("Genre")
        .agg(
            Mean_Survival=("Duration", "mean"),
            Median_Survival=("Duration", "median"),
            Spell_Count=("Duration", "size"),
        )
        .reset_index()
        .sort_values("Mean_Survival", ascending=False, kind="stable")
    )
    variance_summary = (
        df.dropna(subset=["Genre"])
        .groupby("Genre")
        .agg(Rank_Variance=("Rank", "var"), Rank_Mean=("Rank", "mean"), Observations=("Rank", "size"))
        .reset_index()
    )
    summary = summary.merge(variance_summary, on="Genre", how="left")
    summary.to_csv(tables_dir / "genre_momentum_summary.csv", index=False)

    groups = [group["Duration"].to_numpy(dtype=float) for _, group in genre_spells.groupby("Genre")]
    if len(groups) >= 2 and all(len(group) >= 2 for group in groups):
        try:
            kw_stat, kw_p = stats.kruskal(*groups)
            tests.append(
                {
                    "Analysis": "Genre Momentum",
                    "Hypothesis": "H0: Survival distributions are equal across genres",
                    "Test": "Kruskal-Wallis",
                    "Statistic": kw_stat,
                    "PValue": kw_p,
                    "Decision_5pct": "Reject H0" if kw_p < 0.05 else "Fail to reject H0",
                    "Note": "Compares completed and censored spell durations by genre without assuming normality.",
                }
            )
        except ValueError:
            notes.append("Kruskal-Wallis genre test was skipped because all observed genre durations were identical.")
    else:
        notes.append("Kruskal-Wallis genre test was skipped because some genre groups were too small.")

    notes.append("A Cox proportional hazards model is not included because the required survival package is not installed in this environment.")
    return summary, tests, notes


def analyze_markov(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    transitions = df.dropna(subset=["Next_State"]).copy()
    if transitions.empty:
        notes.append("Markov analysis was skipped because the dataset has no next-day chart observations yet.")
        return pd.DataFrame(), tests, notes

    counts = pd.crosstab(transitions["State"], transitions["Next_State"]).reindex(
        index=STATE_ORDER[:-1], columns=STATE_ORDER, fill_value=0
    )
    probs = counts.div(counts.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    full_chain = probs.reindex(index=STATE_ORDER, columns=STATE_ORDER, fill_value=0)
    full_chain.loc["Exit", :] = 0
    full_chain.loc["Exit", "Exit"] = 1.0

    counts.to_csv(tables_dir / "markov_transition_counts.csv")
    full_chain.to_csv(tables_dir / "markov_transition_matrix.csv")

    fig, ax = plt.subplots(figsize=(7, 5))
    image = ax.imshow(full_chain.to_numpy(dtype=float), cmap="Blues", vmin=0, vmax=1)
    ax.set_xticks(range(len(full_chain.columns)))
    ax.set_xticklabels(full_chain.columns, rotation=45, ha="right")
    ax.set_yticks(range(len(full_chain.index)))
    ax.set_yticklabels(full_chain.index)
    ax.set_title("Markov Transition Matrix")
    for i in range(len(full_chain.index)):
        for j in range(len(full_chain.columns)):
            ax.text(j, i, f"{full_chain.iloc[i, j]:.2f}", ha="center", va="center", color="black", fontsize=9)
    fig.colorbar(image, ax=ax, label="Probability")
    fig.tight_layout()
    fig.savefig(figures_dir / "markov_transition_matrix.png", dpi=150)
    plt.close(fig)

    transient = full_chain.loc[STATE_ORDER[:-1], STATE_ORDER[:-1]].to_numpy(dtype=float)
    exit_col = full_chain.loc[STATE_ORDER[:-1], ["Exit"]].to_numpy(dtype=float)
    identity = np.eye(len(transient))
    determinant = np.linalg.det(identity - transient)
    if abs(determinant) > 1e-9:
        fundamental = np.linalg.inv(identity - transient)
        expected_steps = fundamental @ np.ones((len(transient), 1))
        absorption = fundamental @ exit_col
        markov_summary = pd.DataFrame(
            {
                "State": STATE_ORDER[:-1],
                "Expected_Steps_Before_Exit": expected_steps.flatten(),
                "Exit_Absorption_Probability": absorption.flatten(),
            }
        )
        markov_summary.to_csv(tables_dir / "markov_absorption_summary.csv", index=False)
        notes.append("Expected time before exit was computed from the absorbing Markov chain fundamental matrix.")
    else:
        notes.append("Expected time before exit could not be computed because the transient-state matrix is singular.")

    non_exit = probs.loc[STATE_ORDER[:-1], STATE_ORDER[:-1]].copy()
    row_sums = non_exit.sum(axis=1)
    if (row_sums > 0).all():
        conditional_chain = non_exit.div(row_sums, axis=0)
        stationary = np.repeat(1 / len(conditional_chain), len(conditional_chain))
        for _ in range(500):
            stationary = stationary @ conditional_chain.to_numpy(dtype=float)
        stationary_df = pd.DataFrame(
            {"State": conditional_chain.index, "Conditional_Stationary_Probability": stationary}
        )
        stationary_df.to_csv(tables_dir / "markov_conditional_stationary_distribution.csv", index=False)
        notes.append(
            "The full absorbing chain has stationary mass on Exit; a conditional stationary distribution among surviving states was also saved."
        )
    else:
        notes.append("Conditional stationary distribution among surviving states was skipped because at least one state always exits immediately.")

    return full_chain, tests, notes


def analyze_prediction(df: pd.DataFrame, tables_dir: Path, figures_dir: Path) -> tuple[pd.DataFrame, list[dict], list[str]]:
    notes: list[str] = []
    tests: list[dict] = []
    modeling = df.dropna(subset=["Top10_Tomorrow"]).copy()
    if modeling.empty:
        notes.append("Top 10 prediction was skipped because no day has a next-day label yet.")
        return pd.DataFrame(), tests, notes

    date_sequence = sorted(modeling["Date"].drop_duplicates())
    if len(date_sequence) < 4:
        notes.append("Top 10 prediction was skipped because at least 4 chart dates are needed for a train/test split.")
        return modeling, tests, notes

    split_idx = max(2, math.ceil(len(date_sequence) * 0.7))
    train_dates = set(date_sequence[:split_idx])
    test_dates = set(date_sequence[split_idx:])
    train = modeling[modeling["Date"].isin(train_dates)].copy()
    test = modeling[modeling["Date"].isin(test_dates)].copy()

    if train["Top10_Tomorrow"].nunique() < 2 or test.empty or test["Top10_Tomorrow"].nunique() < 2:
        notes.append("Top 10 prediction was skipped because the train/test split does not contain both outcome classes.")
        return modeling, tests, notes

    feature_cols = ["Rank", "Days_On_Chart", "Artist_Past_Appearances", "Volatility_So_Far", "Total_Appearances"]
    train_X = train[feature_cols].copy()
    test_X = test[feature_cols].copy()

    if "Genre" in modeling.columns and modeling["Genre"].notna().any():
        train_X = pd.concat([train_X, pd.get_dummies(train["Genre"], prefix="Genre", dtype=float)], axis=1)
        test_X = pd.concat([test_X, pd.get_dummies(test["Genre"], prefix="Genre", dtype=float)], axis=1)

    train_X, test_X = train_X.align(test_X, join="outer", axis=1, fill_value=0)
    model = LogisticRegression(max_iter=1000, random_state=42)
    model.fit(train_X, train["Top10_Tomorrow"].astype(int))

    probabilities = model.predict_proba(test_X)[:, 1]
    predictions = (probabilities >= 0.5).astype(int)
    metrics = pd.DataFrame(
        [
            {"Metric": "Accuracy", "Value": accuracy_score(test["Top10_Tomorrow"], predictions)},
            {"Metric": "Precision", "Value": precision_score(test["Top10_Tomorrow"], predictions, zero_division=0)},
            {"Metric": "Recall", "Value": recall_score(test["Top10_Tomorrow"], predictions, zero_division=0)},
            {"Metric": "F1", "Value": f1_score(test["Top10_Tomorrow"], predictions, zero_division=0)},
            {"Metric": "ROC_AUC", "Value": roc_auc_score(test["Top10_Tomorrow"], probabilities)},
        ]
    )
    metrics.to_csv(tables_dir / "top10_prediction_metrics.csv", index=False)

    coefficients = pd.DataFrame(
        {"Feature": train_X.columns, "Coefficient": model.coef_[0]}
    ).sort_values("Coefficient", ascending=False, kind="stable")
    coefficients.to_csv(tables_dir / "top10_prediction_coefficients.csv", index=False)

    output_predictions = test[["Date", "Artist", "Song", "Rank", "Top10_Tomorrow"]].copy()
    output_predictions["Predicted_Probability"] = probabilities
    output_predictions["Predicted_Class"] = predictions
    output_predictions.to_csv(tables_dir / "top10_prediction_scored_rows.csv", index=False)

    cm = confusion_matrix(test["Top10_Tomorrow"], predictions)
    fig, ax = plt.subplots(figsize=(5, 4))
    image = ax.imshow(cm, cmap="Greens")
    ax.set_xticks([0, 1])
    ax.set_xticklabels(["Pred 0", "Pred 1"])
    ax.set_yticks([0, 1])
    ax.set_yticklabels(["True 0", "True 1"])
    ax.set_title("Top 10 Prediction Confusion Matrix")
    for i in range(cm.shape[0]):
        for j in range(cm.shape[1]):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center")
    fig.colorbar(image, ax=ax)
    fig.tight_layout()
    fig.savefig(figures_dir / "top10_prediction_confusion_matrix.png", dpi=150)
    plt.close(fig)

    coeff_plot = coefficients.assign(AbsCoeff=coefficients["Coefficient"].abs()).sort_values(
        "AbsCoeff", ascending=False, kind="stable"
    ).head(10)
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(coeff_plot["Feature"], coeff_plot["Coefficient"], color="#e45756")
    ax.set_title("Logistic Regression Coefficients for Staying in Top 10")
    ax.set_xlabel("Coefficient")
    fig.tight_layout()
    fig.savefig(figures_dir / "top10_prediction_coefficients.png", dpi=150)
    plt.close(fig)

    top_rank_coef = coefficients.loc[coefficients["Feature"] == "Rank", "Coefficient"]
    if not top_rank_coef.empty:
        tests.append(
            {
                "Analysis": "Top10 Prediction",
                "Hypothesis": "H0: Today's rank has no effect on Top 10 retention",
                "Test": "Logistic regression coefficient sign",
                "Statistic": float(top_rank_coef.iloc[0]),
                "PValue": np.nan,
                "Decision_5pct": "Interpret coefficient directly",
                "Note": "A more negative coefficient means better current ranks are associated with higher retention odds.",
            }
        )

    days_coef = coefficients.loc[coefficients["Feature"] == "Days_On_Chart", "Coefficient"]
    if not days_coef.empty:
        tests.append(
            {
                "Analysis": "Top10 Prediction",
                "Hypothesis": "H0: Days-on-chart has no effect on Top 10 retention",
                "Test": "Logistic regression coefficient sign",
                "Statistic": float(days_coef.iloc[0]),
                "PValue": np.nan,
                "Decision_5pct": "Interpret coefficient directly",
                "Note": "Positive values indicate longer-running songs are more likely to stay in the Top 10 next day.",
            }
        )

    notes.append("Top 10 prediction uses a chronological train/test split, so it measures next-day forecasting rather than random resampling.")
    return metrics, tests, notes


def save_hypothesis_table(hypothesis_rows: list[dict], tables_dir: Path) -> pd.DataFrame:
    table = pd.DataFrame(hypothesis_rows)
    if table.empty:
        table = pd.DataFrame(
            columns=["Analysis", "Hypothesis", "Test", "Statistic", "PValue", "Decision_5pct", "Note"]
        )
    table.to_csv(tables_dir / "hypothesis_tests.csv", index=False)
    return table


def write_report(
    report_path: Path,
    cleaned_rows: int,
    date_count: int,
    notes: list[str],
    hypothesis_table: pd.DataFrame,
    output_dirs: dict[str, Path],
) -> None:
    figure_names = sorted(path.name for path in output_dirs["figures"].glob("*.png"))
    table_names = sorted(path.name for path in output_dirs["tables"].glob("*.csv"))
    pvalue_lines: list[str] = []
    if not hypothesis_table.empty:
        for row in hypothesis_table.fillna("").itertuples():
            p_value = row.PValue if row.PValue != "" else "NA"
            if isinstance(p_value, float):
                p_value = f"{p_value:.6f}"
            pvalue_lines.append(f"- {row.Analysis}: {row.Test} -> {row.Decision_5pct} (p={p_value})")

    lines = [
        "# Spotify Top 50 Probability Pipeline Report",
        "",
        "## Dataset Snapshot",
        f"- Cleaned chart rows: {cleaned_rows}",
        f"- Unique chart dates: {date_count}",
        "- This pipeline is designed for daily Spotify Top 50 exports stored as one sheet per date.",
        "",
        "## What Was Automated",
        "- Cleaned the workbook into `spotify_data.csv`.",
        "- Computed rank deltas, chart spells, artist concentration metrics, entry rates, Markov transitions, and Top 10 prediction features.",
        "- Generated CSV tables in `output/tables` and figures in `output/figures`.",
        "- Wrote a hypothesis-test summary table for the statistical questions that had enough data.",
        "",
        "## Notes",
    ]
    lines.extend(f"- {note}" for note in notes)
    lines.extend(["", "## Hypothesis Summary"])
    if pvalue_lines:
        lines.extend(pvalue_lines)
    else:
        lines.append("- No formal hypothesis tests were run because the current dataset is too small.")
    lines.extend(["", "## Output Files", "- Tables:"])
    lines.extend(f"  - {name}" for name in table_names)
    lines.append("- Figures:")
    lines.extend(f"  - {name}" for name in figure_names)
    report_path.write_text("\n".join(lines), encoding="utf-8")


def run_pipeline(
    workbook_path: Path,
    output_dir: Path,
    cleaned_csv_path: Path,
    year: int | None = 2026,
    metadata_path: Path | None = None,
) -> dict[str, object]:
    output_dirs = ensure_dirs(output_dir)
    workbook_path = resolve_workbook_path(workbook_path)
    cleaned = clean_workbook(workbook_path, year)
    cleaned.to_csv(cleaned_csv_path, index=False)

    metadata_path = metadata_path or Path("genre_mapping.csv")
    cleaned_with_metadata, notes = merge_genre_metadata(cleaned, metadata_path, output_dirs["base"])
    analyzed, unique_dates = prepare_analysis_frame(cleaned_with_metadata)
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
        cleaned_rows=len(cleaned),
        date_count=len(unique_dates),
        notes=notes,
        hypothesis_table=hypothesis_table,
        output_dirs=output_dirs,
    )

    return {
        "cleaned_rows": len(cleaned),
        "date_count": len(unique_dates),
        "output_dir": str(output_dirs["base"].resolve()),
        "report_path": str((output_dirs["base"] / "report.md").resolve()),
        "cleaned_csv_path": str(cleaned_csv_path.resolve()),
    }
