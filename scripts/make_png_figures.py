#!/usr/bin/env python3
"""Create PNG figures from Four Choice pipeline CSV outputs."""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd


def group_series(df: pd.DataFrame) -> pd.Series:
    if "group" not in df:
        return pd.Series(["All animals"] * len(df), index=df.index)
    group = df["group"].fillna("").astype(str).str.strip()
    if group.replace({"": pd.NA, "N/A": pd.NA, "NA": pd.NA, "nan": pd.NA}).dropna().nunique() <= 1:
        return pd.Series(["All animals"] * len(df), index=df.index)
    return group.replace("", "Unspecified")


def mean_sem_by_group(df: pd.DataFrame, value_cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = df.copy()
    data["_plot_group"] = group_series(data)
    data[value_cols] = data[value_cols].apply(pd.to_numeric, errors="coerce")
    means = data.groupby("_plot_group")[value_cols].mean()
    sems = data.groupby("_plot_group")[value_cols].sem().fillna(0)
    return means, sems


def usable_numeric_cols(df: pd.DataFrame, cols: list[str]) -> list[str]:
    """Return requested columns that contain at least one numeric value."""
    usable = []
    for col in cols:
        if col in df and pd.to_numeric(df[col], errors="coerce").notna().any():
            usable.append(col)
    return usable


def grouped_bars(ax, means: pd.DataFrame, sems: pd.DataFrame, colors: list[str], ylabel: str, title: str) -> None:
    groups = list(means.index)
    cols = list(means.columns)
    width = 0.8 / max(len(groups), 1)
    x = range(len(cols))
    for gi, group in enumerate(groups):
        offset = (gi - (len(groups) - 1) / 2) * width
        positions = [i + offset for i in x]
        ax.bar(positions, means.loc[group, cols], width=width, yerr=sems.loc[group, cols], capsize=4, label=group, color=colors[gi % len(colors)], alpha=0.82)
    ax.set_xticks(list(x), [c.replace("_", " ") for c in cols], rotation=20, ha="right")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    if len(groups) > 1:
        ax.legend(frameon=False)


def pca_axis_label(tables: Path, component: str) -> str:
    variance_path = tables / "pca_variance.csv"
    if not variance_path.exists():
        return component
    variance = pd.read_csv(variance_path)
    if not {"component", "explained_variance_ratio"}.issubset(variance.columns):
        return component
    row = variance[variance["component"] == component]
    if row.empty:
        return component
    percent = pd.to_numeric(row["explained_variance_ratio"], errors="coerce").iloc[0] * 100
    if pd.isna(percent):
        return component
    return f"{component} ({percent:.1f}%)"


def plot_ttc(animals: pd.DataFrame, figures: Path) -> None:
    cols = ["TTC_in_discrimination", "TTC_in_reversal"]
    if usable_numeric_cols(animals, cols) != cols:
        return
    values = animals[cols].apply(pd.to_numeric, errors="coerce")
    fig, ax = plt.subplots(figsize=(6.5, 4.5))
    x = [0, 1]
    means = values.mean()
    sems = values.sem().fillna(0)
    ax.bar(x, means, yerr=sems, color=["#4C78A8", "#F58518"], alpha=0.78, capsize=5)
    for _, row in values.iterrows():
        if row.notna().all():
            ax.plot(x, row.values, color="#666666", alpha=0.45, linewidth=1)
            ax.scatter(x, row.values, color="#222222", s=24, zorder=3)
    ax.set_xticks(x, ["Discrimination", "Reversal"])
    ax.set_ylabel("Trials to criterion")
    ax.set_title("Trials to criterion by animal")
    for i, value in enumerate(means):
        ax.text(i, value, f"mean={value:.1f}", ha="center", va="bottom", fontsize=8)
    fig.tight_layout()
    fig.savefig(figures / "trials_to_criterion_by_animal.png", dpi=300)
    plt.close(fig)


def plot_trials_to_criterion_mean_sem(animals: pd.DataFrame, figures: Path) -> None:
    cols = ["TTC_in_discrimination", "TTC_in_reversal"]
    if usable_numeric_cols(animals, cols) != cols:
        return
    means, sems = mean_sem_by_group(animals, cols)
    fig, ax = plt.subplots(figsize=(7, 4.8))
    grouped_bars(ax, means, sems, ["#4C78A8", "#F58518", "#54A24B"], "Trials to criterion", "Trials to criterion")
    fig.tight_layout()
    fig.savefig(figures / "trials_to_criterion_mean_sem.png", dpi=300)
    plt.close(fig)


def odor_roles(row: pd.Series) -> dict[str, str]:
    rewarded = row.get("rewarded_odor", "B")
    if rewarded in {"A", "B", "C", "D"}:
        roles = {"A": "Previously rewarded", "B": "Rewarded", "C": "Irrelevant", "D": "Novel"}
    else:
        roles = {"O1": "Previously rewarded", "O2": "Rewarded", "O3": "Irrelevant", "O4": "Novel"}
    if rewarded not in roles:
        roles[rewarded] = "Rewarded"
    return roles


def plot_four_choice_reversal_arena(tables: Path, figures: Path) -> None:
    trial_path = tables / "trial_level.csv"
    if not trial_path.exists():
        return
    trials = pd.read_csv(trial_path)
    reversal = trials[trials["phase"] == "reversal"].copy()
    if reversal.empty:
        return
    row = reversal.iloc[0]
    order = str(row.get("bowl_order", "")).replace(" ", "")
    if len(order) != 4:
        return
    quadrant_order = ["NW", "NE", "SW", "SE"]
    locations = {
        "NW": (0.25, 0.72),
        "NE": (0.75, 0.72),
        "SW": (0.25, 0.28),
        "SE": (0.75, 0.28),
    }
    colors = {
        "Rewarded": "#54A24B",
        "Previously rewarded": "#4C78A8",
        "Irrelevant": "#B279A2",
        "Novel": "#F58518",
    }
    roles = odor_roles(row)
    fig, ax = plt.subplots(figsize=(6.8, 5.4))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.set_title("Four-choice reversal arena", fontsize=14, weight="bold", pad=12)
    ax.add_patch(plt.Rectangle((0.08, 0.08), 0.84, 0.78, fill=False, linewidth=2, edgecolor="#222222"))
    ax.plot([0.5, 0.5], [0.08, 0.86], color="#222222", linewidth=1.5)
    ax.plot([0.08, 0.92], [0.47, 0.47], color="#222222", linewidth=1.5)
    ax.add_patch(plt.Circle((0.5, 0.47), 0.095, fill=False, linewidth=1.8, edgecolor="#555555"))
    ax.text(0.5, 0.47, "Start\ncylinder", ha="center", va="center", fontsize=10)
    for quadrant, odor in zip(quadrant_order, order):
        role = roles.get(odor, "")
        x, y = locations[quadrant]
        ax.add_patch(plt.Circle((x, y), 0.075, color=colors.get(role, "#CCCCCC"), alpha=0.85))
        ax.text(x, y + 0.005, odor, ha="center", va="center", fontsize=16, weight="bold", color="white")
        ax.text(x, y - 0.12, f"{quadrant}\n{role}", ha="center", va="top", fontsize=9)
    ax.text(0.5, 0.02, "Each trial maps entry digits 1-4 to NW, NE, SW, SE, then back to odor by ramekin order.", ha="center", fontsize=9)
    handles = [
        plt.Line2D([0], [0], marker="o", color="w", label=role, markerfacecolor=color, markersize=10)
        for role, color in colors.items()
    ]
    ax.legend(handles=handles, loc="upper center", bbox_to_anchor=(0.5, -0.06), ncol=2, frameon=False, fontsize=9)
    fig.tight_layout()
    fig.savefig(figures / "four_choice_reversal_arena.png", dpi=300, bbox_inches="tight")
    plt.close(fig)


def plot_reversal_error_types(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(
        animals,
        [
            "Total_errors_in_reversal",
            "Perseverative_errors_in_reversal",
            "Regressive_errors_in_reversal",
            "rev_irrelevant_errors",
            "rev_novel_errors",
            "rev_omissions",
        ],
    )
    if not cols:
        return
    means, sems = mean_sem_by_group(animals, cols)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    grouped_bars(ax, means, sems, ["#72B7B2", "#E45756", "#4C78A8"], "Errors", "Reversal error types")
    fig.tight_layout()
    fig.savefig(figures / "reversal_error_types.png", dpi=300)
    plt.close(fig)


def plot_reversal_choice_percent(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["rev_A_choice_percent_trials", "rev_B_choice_percent_trials", "rev_C_choice_percent_trials", "rev_D_choice_percent_trials"])
    if not cols:
        return
    labels = {
        "rev_A_choice_percent_trials": "Previously rewarded",
        "rev_B_choice_percent_trials": "Rewarded",
        "rev_C_choice_percent_trials": "Irrelevant",
        "rev_D_choice_percent_trials": "Novel",
    }
    means, sems = mean_sem_by_group(animals, cols)
    means = means.rename(columns=labels)
    sems = sems.rename(columns=labels)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    grouped_bars(ax, means, sems, ["#4C78A8", "#F58518", "#54A24B"], "% of reversal trials", "Reversal choices by odor")
    ax.axhline(25, color="#444444", linestyle=":", linewidth=1)
    ax.text(len(means.columns) - 0.15, 25.8, "chance", ha="right", va="bottom", fontsize=8)
    fig.tight_layout()
    fig.savefig(figures / "reversal_choice_percent.png", dpi=300)
    plt.close(fig)


def plot_error_types(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["Perseverative_errors_in_reversal", "Regressive_errors_in_reversal", "O1_errors", "O3_errors", "O4_errors"])
    if not cols:
        return
    means = animals[cols].apply(pd.to_numeric, errors="coerce").mean().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.bar(range(len(means)), means.values, color="#54A24B", alpha=0.82)
    ax.set_xticks(range(len(means)), [c.replace("_", " ") for c in means.index], rotation=25, ha="right")
    ax.set_ylabel("Mean errors")
    ax.set_title("Reversal error-type summary")
    for i, value in enumerate(means.values):
        ax.text(i, value, f"{value:.1f}", ha="center", va="bottom", fontsize=8)
    fig.tight_layout()
    fig.savefig(figures / "reversal_error_types_mean.png", dpi=300)
    plt.close(fig)


def plot_entries(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["recall_mean_entries", "rev_mean_entries"])
    if not cols:
        return
    data = animals[cols].apply(pd.to_numeric, errors="coerce")
    fig, ax = plt.subplots(figsize=(6.5, 4.5))
    ax.boxplot([data[c].dropna() for c in cols], labels=[c.replace("_", " ") for c in cols], showmeans=True)
    ax.set_ylabel("Entries")
    ax.set_title("Entry / locomotor readout")
    fig.tight_layout()
    fig.savefig(figures / "entry_locomotor_summary.png", dpi=300)
    plt.close(fig)


def plot_reversal_latency(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["rev_mean_latency_sec", "rev_mean_latency_correct_sec", "rev_mean_latency_incorrect_sec"])
    if not cols:
        return
    labels = {
        "rev_mean_latency_sec": "All trials",
        "rev_mean_latency_correct_sec": "Correct",
        "rev_mean_latency_incorrect_sec": "Incorrect",
    }
    means, sems = mean_sem_by_group(animals, cols)
    means = means.rename(columns=labels)
    sems = sems.rename(columns=labels)
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    grouped_bars(ax, means, sems, ["#4C78A8", "#F58518", "#54A24B"], "Latency to dig (s)", "Reversal latency")
    fig.tight_layout()
    fig.savefig(figures / "reversal_latency.png", dpi=300)
    plt.close(fig)


def plot_reversal_entries(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["rev_mean_entries", "rev_mean_entries_correct", "rev_mean_entries_incorrect"])
    if not cols:
        return
    labels = {
        "rev_mean_entries": "All trials",
        "rev_mean_entries_correct": "Correct",
        "rev_mean_entries_incorrect": "Incorrect",
    }
    means, sems = mean_sem_by_group(animals, cols)
    means = means.rename(columns=labels)
    sems = sems.rename(columns=labels)
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    grouped_bars(ax, means, sems, ["#4C78A8", "#F58518", "#54A24B"], "Entries per trial", "Reversal entries")
    fig.tight_layout()
    fig.savefig(figures / "reversal_entries.png", dpi=300)
    plt.close(fig)


def plot_reversal_entry_percent(animals: pd.DataFrame, figures: Path) -> None:
    cols = usable_numeric_cols(animals, ["rev_A_entry_percent_total", "rev_B_entry_percent_total", "rev_C_entry_percent_total", "rev_D_entry_percent_total"])
    if not cols:
        return
    labels = {
        "rev_A_entry_percent_total": "Previously rewarded",
        "rev_B_entry_percent_total": "Rewarded",
        "rev_C_entry_percent_total": "Irrelevant",
        "rev_D_entry_percent_total": "Novel",
    }
    means, sems = mean_sem_by_group(animals, cols)
    means = means.rename(columns=labels)
    sems = sems.rename(columns=labels)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    grouped_bars(ax, means, sems, ["#4C78A8", "#F58518", "#54A24B"], "% of total entries", "Entries by odor quadrant")
    ax.axhline(25, color="#444444", linestyle=":", linewidth=1)
    ax.text(len(means.columns) - 0.15, 25.8, "chance", ha="right", va="bottom", fontsize=8)
    fig.tight_layout()
    fig.savefig(figures / "reversal_entry_percent.png", dpi=300)
    plt.close(fig)


def plot_pca(tables: Path, figures: Path) -> None:
    pca_path = tables / "pca_scores.csv"
    if not pca_path.exists():
        return
    scores = pd.read_csv(pca_path)
    if not {"PC1", "PC2"}.issubset(scores.columns):
        return
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    ax.scatter(scores["PC1"], scores["PC2"], color="#B279A2", s=48)
    for _, row in scores.iterrows():
        ax.text(row["PC1"], row["PC2"], str(row["animal_id"]), fontsize=8, ha="left", va="bottom")
    ax.axhline(0, color="#AAAAAA", linewidth=0.8)
    ax.axvline(0, color="#AAAAAA", linewidth=0.8)
    ax.set_xlabel(pca_axis_label(tables, "PC1"))
    ax.set_ylabel(pca_axis_label(tables, "PC2"))
    ax.set_title("First PCA from Four Choice outputs")
    fig.tight_layout()
    fig.savefig(figures / "pca_scores.png", dpi=300)
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--outdir", type=Path, default=Path("four_choice_tabs_pipeline/outputs/example_analysis"))
    args = parser.parse_args()

    tables = args.outdir / "tables"
    figures = args.outdir / "figures"
    figures.mkdir(parents=True, exist_ok=True)
    animals = pd.read_csv(tables / "animal_level_outputs.csv")
    plot_ttc(animals, figures)
    plot_four_choice_reversal_arena(tables, figures)
    plot_trials_to_criterion_mean_sem(animals, figures)
    plot_error_types(animals, figures)
    plot_reversal_error_types(animals, figures)
    plot_reversal_choice_percent(animals, figures)
    plot_entries(animals, figures)
    plot_reversal_latency(animals, figures)
    plot_reversal_entries(animals, figures)
    plot_reversal_entry_percent(animals, figures)
    plot_pca(tables, figures)
    print(f"Wrote PNG figures to {figures.resolve()}")


if __name__ == "__main__":
    main()
