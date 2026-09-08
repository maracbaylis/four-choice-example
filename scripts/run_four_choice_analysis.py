#!/usr/bin/env python3
"""Run the ARC Four Choice analysis pipeline on filled TABS workbooks."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from four_choice_core import add_metadata_context, autosize_excel, metadata_qc, parse_tabs_workbook, pca_table, submission_qc, summarize_animals, summarize_sessions  # noqa: E402


DEFAULT_OUT = Path("four_choice_tabs_pipeline/outputs/example_analysis")
QC_COLUMNS = ["animal_id", "source_file", "field", "qc_flag", "value", "message"]


def workbook_paths(inputs: list[Path]) -> list[Path]:
    paths: list[Path] = []
    for item in inputs:
        if item.is_dir():
            paths.extend(sorted(item.glob("*.xlsx")))
        elif item.suffix.lower() == ".xlsx":
            paths.append(item)
    return [p for p in paths if not p.name.startswith("~$")]


def concat_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    usable = [frame.dropna(axis=1, how="all") for frame in frames if not frame.empty]
    return pd.concat(usable, ignore_index=True) if usable else pd.DataFrame()


def use_workbook_names_for_source_files(frame: pd.DataFrame) -> pd.DataFrame:
    """Keep source-file labels portable in saved outputs."""
    if frame.empty or "source_file" not in frame.columns:
        return frame
    frame = frame.copy()
    frame["source_file"] = frame["source_file"].map(lambda value: Path(str(value)).name if pd.notna(value) else "")
    return frame


def build_qc_table(metadata: pd.DataFrame, trials: pd.DataFrame, weights: pd.DataFrame) -> pd.DataFrame:
    qc = concat_frames([metadata_qc(metadata), submission_qc(metadata, trials, weights)])
    if qc.empty:
        return pd.DataFrame(columns=QC_COLUMNS)
    return qc.reindex(columns=QC_COLUMNS)


def print_qc_summary(qc: pd.DataFrame, max_messages: int = 20) -> None:
    if qc.empty:
        print("QC: no metadata or submission issues flagged.")
        return

    print(f"QC: {len(qc)} issue(s) flagged. Review tables/metadata_qc.csv for the full report.")
    for _, row in qc.head(max_messages).iterrows():
        animal_id = row.get("animal_id", "unknown_animal")
        source_file = row.get("source_file", "")
        field = row.get("field", "")
        flag = row.get("qc_flag", "")
        message = row.get("message", "")
        where = f" [{source_file}]" if isinstance(source_file, str) and source_file else ""
        print(f"  - {animal_id}{where}: {field} ({flag}) - {message}")
    remaining = len(qc) - max_messages
    if remaining > 0:
        print(f"  ... {remaining} more issue(s) in tables/metadata_qc.csv")


def parse_all(paths: list[Path]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    trials = []
    metadata = []
    weights = []
    acclim = []
    for path in paths:
        trial_df, meta_df, weight_df, acclim_df = parse_tabs_workbook(path)
        if not trial_df.empty:
            trials.append(trial_df)
        metadata.append(meta_df)
        if not weight_df.empty:
            weights.append(weight_df)
        if not acclim_df.empty:
            acclim.append(acclim_df)
    return tuple(
        use_workbook_names_for_source_files(frame)
        for frame in (concat_frames(trials), concat_frames(metadata), concat_frames(weights), concat_frames(acclim))
    )


def write_excel_output(out_xlsx: Path, trials: pd.DataFrame, sessions: pd.DataFrame, animals: pd.DataFrame, weights: pd.DataFrame, acclim: pd.DataFrame, qc: pd.DataFrame, pca_scores: pd.DataFrame, pca_loadings: pd.DataFrame, pca_variance: pd.DataFrame) -> None:
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        trials.to_excel(writer, sheet_name="trial_level", index=False)
        sessions.to_excel(writer, sheet_name="session_level", index=False)
        animals.to_excel(writer, sheet_name="animal_level_outputs", index=False)
        if not weights.empty:
            weights.to_excel(writer, sheet_name="weight_long", index=False)
        if not acclim.empty:
            acclim.to_excel(writer, sheet_name="acclim_long", index=False)
        qc.to_excel(writer, sheet_name="metadata_qc", index=False)
        core_cols = [
            "animal_id",
            "lab_source",
            "experiment_cohort",
            "group",
            "sex",
            "genotype",
            "TTC_in_discrimination",
            "Total_errors_in_discrimination",
            "TTC_in_reversal",
            "Total_errors_in_reversal",
            "Perseverative_errors_in_reversal",
            "Regressive_errors_in_reversal",
            "O1_errors",
            "O3_errors",
            "O4_errors",
            "rev_A_choice_percent_trials",
            "rev_B_choice_percent_trials",
            "rev_C_choice_percent_trials",
            "rev_D_choice_percent_trials",
            "rev_A_entry_percent_total",
            "rev_B_entry_percent_total",
            "rev_C_entry_percent_total",
            "rev_D_entry_percent_total",
            "rev_mean_latency_correct_sec",
            "rev_mean_latency_incorrect_sec",
            "rev_mean_entries_correct",
            "rev_mean_entries_incorrect",
        ]
        animals[[c for c in core_cols if c in animals.columns]].to_excel(writer, sheet_name="core_variables", index=False)
        make_prism_tables(animals).to_excel(writer, sheet_name="prism_long", index=False)
        if not pca_scores.empty:
            pca_scores.to_excel(writer, sheet_name="pca_scores", index=False)
            pca_loadings.to_excel(writer, sheet_name="pca_loadings", index=False)
            pca_variance.to_excel(writer, sheet_name="pca_variance", index=False)
    autosize_excel(out_xlsx)


def make_prism_tables(animals: pd.DataFrame) -> pd.DataFrame:
    id_cols = [c for c in ["animal_id", "lab_source", "experiment_cohort", "group", "sex", "genotype"] if c in animals.columns]
    value_cols = [
        c
        for c in [
            "TTC_in_discrimination",
            "Total_errors_in_discrimination",
            "TTC_in_reversal",
            "Total_errors_in_reversal",
            "Perseverative_errors_in_reversal",
            "Regressive_errors_in_reversal",
            "O1_errors",
            "O3_errors",
            "O4_errors",
            "recall_mean_entries",
            "rev_mean_entries",
            "rev_total_entries",
            "rev_mean_latency_sec",
            "rev_mean_latency_correct_sec",
            "rev_mean_latency_incorrect_sec",
            "rev_mean_entries_correct",
            "rev_mean_entries_incorrect",
            "rev_A_choice_percent_trials",
            "rev_B_choice_percent_trials",
            "rev_C_choice_percent_trials",
            "rev_D_choice_percent_trials",
            "rev_A_entry_percent_total",
            "rev_B_entry_percent_total",
            "rev_C_entry_percent_total",
            "rev_D_entry_percent_total",
        ]
        if c in animals.columns
    ]
    return animals.melt(id_vars=id_cols, value_vars=value_cols, var_name="variable", value_name="value")


def save_tables(outdir: Path, trials: pd.DataFrame, sessions: pd.DataFrame, animals: pd.DataFrame, weights: pd.DataFrame, acclim: pd.DataFrame, qc: pd.DataFrame, pca_scores: pd.DataFrame, pca_loadings: pd.DataFrame, pca_variance: pd.DataFrame) -> None:
    tables = outdir / "tables"
    tables.mkdir(parents=True, exist_ok=True)
    trials.to_csv(tables / "trial_level.csv", index=False)
    sessions.to_csv(tables / "session_level.csv", index=False)
    animals.to_csv(tables / "animal_level_outputs.csv", index=False)
    if not weights.empty:
        weights.to_csv(tables / "weight_long.csv", index=False)
    if not acclim.empty:
        acclim.to_csv(tables / "acclim_long.csv", index=False)
    qc.to_csv(tables / "metadata_qc.csv", index=False)
    make_prism_tables(animals).to_csv(tables / "prism_long.csv", index=False)
    if not pca_scores.empty:
        pca_scores.to_csv(tables / "pca_scores.csv", index=False)
        pca_loadings.to_csv(tables / "pca_loadings.csv", index=False)
        pca_variance.to_csv(tables / "pca_variance.csv", index=False)


def svg_header(width: int, height: int) -> list[str]:
    return [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        '<style>text{font-family:Arial,Helvetica,sans-serif;font-size:12px;fill:#222}.title{font-size:16px;font-weight:bold}.axis{stroke:#333;stroke-width:1}.grid{stroke:#ddd;stroke-width:1}.label{font-size:11px}</style>',
    ]


def save_svg(lines: list[str], path: Path) -> None:
    path.write_text("\n".join(lines + ["</svg>\n"]))


def scale_y(value: float, ymax: float, top: int = 50, bottom: int = 310) -> float:
    if ymax <= 0:
        return bottom
    return bottom - (value / ymax) * (bottom - top)


def plot_ttc(animals: pd.DataFrame, outdir: Path) -> None:
    cols = ["TTC_in_discrimination", "TTC_in_reversal"]
    if not all(col in animals for col in cols):
        return
    means = [animals[col].mean() for col in cols]
    sems = [animals[col].sem() if len(animals[col].dropna()) > 1 else 0 for col in cols]
    ymax = max([v for v in means + sems if pd.notna(v)] + [1]) * 1.25
    xs = [210, 390]
    width = 72
    lines = svg_header(560, 360)
    lines.append('<text x="280" y="24" text-anchor="middle" class="title">Trials to criterion by animal</text>')
    lines.append('<line x1="80" y1="310" x2="500" y2="310" class="axis"/>')
    lines.append('<line x1="80" y1="50" x2="80" y2="310" class="axis"/>')
    for i, mean in enumerate(means):
        if pd.isna(mean):
            continue
        y = scale_y(float(mean), ymax)
        h = 310 - y
        color = ["#4C78A8", "#F58518"][i]
        lines.append(f'<rect x="{xs[i]-width/2}" y="{y:.1f}" width="{width}" height="{h:.1f}" fill="{color}" opacity="0.75"/>')
        err = float(sems[i] or 0)
        yerr = scale_y(float(mean + err), ymax)
        lines.append(f'<line x1="{xs[i]}" y1="{yerr:.1f}" x2="{xs[i]}" y2="{y:.1f}" stroke="#222" stroke-width="1.4"/>')
        lines.append(f'<line x1="{xs[i]-18}" y1="{yerr:.1f}" x2="{xs[i]+18}" y2="{yerr:.1f}" stroke="#222" stroke-width="1.4"/>')
        lines.append(f'<text x="{xs[i]}" y="{y-6:.1f}" text-anchor="middle" class="label">mean={mean:.1f}</text>')
    for _, row in animals.iterrows():
        y0, y1 = row.get(cols[0]), row.get(cols[1])
        if pd.notna(y0) and pd.notna(y1):
            sy0, sy1 = scale_y(float(y0), ymax), scale_y(float(y1), ymax)
            lines.append(f'<line x1="{xs[0]}" y1="{sy0:.1f}" x2="{xs[1]}" y2="{sy1:.1f}" stroke="#666" stroke-width="1" opacity="0.45"/>')
            lines.append(f'<circle cx="{xs[0]}" cy="{sy0:.1f}" r="3" fill="#222"/>')
            lines.append(f'<circle cx="{xs[1]}" cy="{sy1:.1f}" r="3" fill="#222"/>')
    lines.append('<text x="210" y="335" text-anchor="middle">Discrimination</text>')
    lines.append('<text x="390" y="335" text-anchor="middle">Reversal</text>')
    lines.append('<text x="25" y="180" text-anchor="middle" transform="rotate(-90 25 180)">Trials to criterion</text>')
    save_svg(lines, outdir / "trials_to_criterion_by_animal.svg")


def plot_error_types(animals: pd.DataFrame, outdir: Path) -> None:
    cols = [c for c in ["Perseverative_errors_in_reversal", "Regressive_errors_in_reversal", "O1_errors", "O3_errors", "O4_errors"] if c in animals]
    if not cols:
        return
    means = animals[cols].mean().sort_values(ascending=False)
    ymax = max([float(v) for v in means.values if pd.notna(v)] + [1]) * 1.25
    lines = svg_header(720, 380)
    lines.append('<text x="360" y="24" text-anchor="middle" class="title">Reversal error-type summary</text>')
    lines.append('<line x1="80" y1="310" x2="680" y2="310" class="axis"/>')
    lines.append('<line x1="80" y1="50" x2="80" y2="310" class="axis"/>')
    bar_w = min(70, 500 / max(len(means), 1))
    for i, (label, value) in enumerate(means.items()):
        x = 120 + i * (560 / max(len(means), 1))
        y = scale_y(float(value), ymax)
        lines.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{310-y:.1f}" fill="#54A24B" opacity="0.8"/>')
        lines.append(f'<text x="{x+bar_w/2:.1f}" y="{y-6:.1f}" text-anchor="middle" class="label">{value:.1f}</text>')
        short = label.replace("_", " ")
        lines.append(f'<text x="{x+bar_w/2:.1f}" y="330" text-anchor="end" transform="rotate(-25 {x+bar_w/2:.1f} 330)" class="label">{short}</text>')
    lines.append('<text x="25" y="180" text-anchor="middle" transform="rotate(-90 25 180)">Mean errors</text>')
    save_svg(lines, outdir / "reversal_error_types_mean.svg")


def plot_recall_entries(animals: pd.DataFrame, outdir: Path) -> None:
    cols = [c for c in ["recall_mean_entries", "rev_mean_entries"] if c in animals]
    if not cols:
        return
    data = animals[cols].apply(pd.to_numeric, errors="coerce")
    ymax = max([float(v) for v in data.max().dropna().values] + [1]) * 1.25
    lines = svg_header(560, 360)
    lines.append('<text x="280" y="24" text-anchor="middle" class="title">Entry / locomotor readout</text>')
    lines.append('<line x1="80" y1="310" x2="500" y2="310" class="axis"/>')
    lines.append('<line x1="80" y1="50" x2="80" y2="310" class="axis"/>')
    xs = [210, 390]
    for i, col in enumerate(cols):
        vals = data[col].dropna().sort_values()
        if vals.empty:
            continue
        q1, med, q3 = vals.quantile([0.25, 0.5, 0.75])
        lo, hi = vals.min(), vals.max()
        x = xs[i]
        yq1, ymed, yq3 = scale_y(q1, ymax), scale_y(med, ymax), scale_y(q3, ymax)
        ylo, yhi = scale_y(lo, ymax), scale_y(hi, ymax)
        lines.append(f'<line x1="{x}" y1="{yhi:.1f}" x2="{x}" y2="{ylo:.1f}" stroke="#333"/>')
        lines.append(f'<rect x="{x-35}" y="{yq3:.1f}" width="70" height="{yq1-yq3:.1f}" fill="#ECA82C" opacity="0.65" stroke="#333"/>')
        lines.append(f'<line x1="{x-35}" y1="{ymed:.1f}" x2="{x+35}" y2="{ymed:.1f}" stroke="#333" stroke-width="2"/>')
        lines.append(f'<text x="{x}" y="335" text-anchor="middle">{col.replace("_", " ")}</text>')
    lines.append('<text x="25" y="180" text-anchor="middle" transform="rotate(-90 25 180)">Entries</text>')
    save_svg(lines, outdir / "entry_locomotor_summary.svg")


def pca_axis_label(pca_variance: pd.DataFrame, component: str) -> str:
    if pca_variance.empty or not {"component", "explained_variance_ratio"}.issubset(pca_variance.columns):
        return component
    row = pca_variance[pca_variance["component"] == component]
    if row.empty:
        return component
    percent = pd.to_numeric(row["explained_variance_ratio"], errors="coerce").iloc[0] * 100
    if pd.isna(percent):
        return component
    return f"{component} ({percent:.1f}%)"


def plot_pca(pca_scores: pd.DataFrame, pca_variance: pd.DataFrame, outdir: Path) -> None:
    if pca_scores.empty or not {"PC1", "PC2"}.issubset(pca_scores.columns):
        return
    xvals = pca_scores["PC1"].astype(float)
    yvals = pca_scores["PC2"].astype(float)
    xmin, xmax = xvals.min(), xvals.max()
    ymin, ymax = yvals.min(), yvals.max()
    xpad = (xmax - xmin or 1) * 0.2
    ypad = (ymax - ymin or 1) * 0.2
    xmin, xmax, ymin, ymax = xmin - xpad, xmax + xpad, ymin - ypad, ymax + ypad
    def sx(v: float) -> float:
        return 70 + (v - xmin) / (xmax - xmin or 1) * 420
    def sy(v: float) -> float:
        return 310 - (v - ymin) / (ymax - ymin or 1) * 250

    lines = svg_header(560, 360)
    lines.append('<text x="280" y="24" text-anchor="middle" class="title">First PCA from Four Choice outputs</text>')
    lines.append('<line x1="70" y1="310" x2="500" y2="310" class="axis"/>')
    lines.append('<line x1="70" y1="50" x2="70" y2="310" class="axis"/>')
    if xmin < 0 < xmax:
        x0 = sx(0)
        lines.append(f'<line x1="{x0:.1f}" y1="50" x2="{x0:.1f}" y2="310" stroke="#aaa"/>')
    if ymin < 0 < ymax:
        y0 = sy(0)
        lines.append(f'<line x1="70" y1="{y0:.1f}" x2="500" y2="{y0:.1f}" stroke="#aaa"/>')
    for _, row in pca_scores.iterrows():
        x, y = sx(float(row["PC1"])), sy(float(row["PC2"]))
        lines.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="#B279A2"/>')
        lines.append(f'<text x="{x+6:.1f}" y="{y-6:.1f}" class="label">{row["animal_id"]}</text>')
    lines.append(f'<text x="285" y="345" text-anchor="middle">{pca_axis_label(pca_variance, "PC1")}</text>')
    lines.append(f'<text x="25" y="180" text-anchor="middle" transform="rotate(-90 25 180)">{pca_axis_label(pca_variance, "PC2")}</text>')
    save_svg(lines, outdir / "pca_scores.svg")


def write_readme(outdir: Path, paths: list[Path], trials: pd.DataFrame, sessions: pd.DataFrame, animals: pd.DataFrame, weights: pd.DataFrame, acclim: pd.DataFrame, qc: pd.DataFrame) -> None:
    text = f"""# Four Choice TABS Analysis Run

Input workbooks: {len(paths)}
Trial rows: {len(trials)}
Session rows: {len(sessions)}
Animal rows: {len(animals)}
Weight rows: {len(weights)}
Acclimation rows: {len(acclim)}
Metadata QC flags: {len(qc)}

Main outputs:
- `four_choice_master_outputs.xlsx`
- `tables/trial_level.csv`
- `tables/session_level.csv`
- `tables/animal_level_outputs.csv`
- `tables/weight_long.csv`
- `tables/acclim_long.csv`
- `tables/metadata_qc.csv`
- `tables/prism_long.csv`
- `tables/pca_scores.csv`, `tables/pca_loadings.csv`, and
  `tables/pca_variance.csv` when at least two animals and two usable variables
  are present
- `figures/` contains any plots that can be made from non-empty variables

Core variables include TTC and total errors for discrimination/reversal,
perseverative/regressive reversal errors, O1/O3/O4 reversal errors, omissions,
latency, entry/locomotor summaries, odor-choice percentages, and odor-entry
percentages derived by cross-referencing entry sequence quadrants against each
trial's ramekin order.

The left side of the CSV outputs carries metadata context. The first ten parsed
metadata fields are treated as primary grouping and sorting variables rather
than behavior scores. Additional metadata fields are retained as possible
moderator or noise variables for later cross-lab review.

Some figure types require recall or reversal rows. If those sheets are empty in
the input workbooks, the corresponding figures are skipped.
"""
    (outdir / "README.md").write_text(text)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", type=Path, help="Filled TABS workbook(s) or directories containing .xlsx files.")
    parser.add_argument("--outdir", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()

    paths = workbook_paths(args.inputs)
    if not paths:
        raise SystemExit("No .xlsx TABS workbooks found.")

    trials, metadata, weights, acclim = parse_all(paths)
    trials = add_metadata_context(trials, metadata)
    weights = add_metadata_context(weights, metadata)
    acclim = add_metadata_context(acclim, metadata)
    sessions = add_metadata_context(summarize_sessions(trials), metadata)
    animals = summarize_animals(trials, metadata)
    qc = build_qc_table(metadata, trials, weights)
    pca_scores, pca_loadings, pca_variance = pca_table(animals)

    args.outdir.mkdir(parents=True, exist_ok=True)
    figures = args.outdir / "figures"
    figures.mkdir(exist_ok=True)
    save_tables(args.outdir, trials, sessions, animals, weights, acclim, qc, pca_scores, pca_loadings, pca_variance)
    write_excel_output(args.outdir / "four_choice_master_outputs.xlsx", trials, sessions, animals, weights, acclim, qc, pca_scores, pca_loadings, pca_variance)
    plot_ttc(animals, figures)
    plot_error_types(animals, figures)
    plot_recall_entries(animals, figures)
    plot_pca(pca_scores, pca_variance, figures)
    write_readme(args.outdir, paths, trials, sessions, animals, weights, acclim, qc)

    print(f"Parsed {len(paths)} workbook(s).")
    print(f"Wrote {len(trials)} trial rows, {len(sessions)} session rows, {len(animals)} animal rows, {len(weights)} weight rows, and {len(acclim)} acclimation rows.")
    print_qc_summary(qc)
    print(f"Output directory: {args.outdir.resolve()}")


if __name__ == "__main__":
    main()
