"""Public API for running the Four Choice TABS workflow."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPT_DIR = REPO_ROOT / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from four_choice_core import add_metadata_context, pca_table, summarize_animals, summarize_sessions  # noqa: E402
from run_four_choice_analysis import (  # noqa: E402
    build_qc_table,
    parse_all,
    save_tables,
    workbook_paths,
    write_excel_output,
    write_readme,
)


@dataclass(frozen=True)
class FourChoiceRunResult:
    """Tables and output paths from one pipeline run."""

    outdir: Path
    input_workbooks: list[Path]
    trials: pd.DataFrame
    metadata: pd.DataFrame
    weights: pd.DataFrame
    acclim: pd.DataFrame
    sessions: pd.DataFrame
    animals: pd.DataFrame
    qc: pd.DataFrame
    pca_scores: pd.DataFrame
    pca_loadings: pd.DataFrame
    pca_variance: pd.DataFrame


def example_input_dir() -> Path:
    """Return the bundled complete example workbook directory."""
    return REPO_ROOT / "example_data"


def run_pipeline(inputs: str | Path | list[str | Path], outdir: str | Path, make_png: bool = True) -> FourChoiceRunResult:
    """Run the Four Choice workflow on filled TABS workbook(s).

    Parameters
    ----------
    inputs
        A workbook path, a directory of `.xlsx` workbooks, or a list of either.
    outdir
        Output directory for tables, figures, run notes, and the master workbook.
    make_png
        Whether to generate PNG figures in addition to CSV/XLSX outputs.
    """
    if isinstance(inputs, (str, Path)):
        input_items = [Path(inputs)]
    else:
        input_items = [Path(item) for item in inputs]
    paths = workbook_paths(input_items)
    if not paths:
        raise FileNotFoundError("No .xlsx TABS workbooks found in the provided inputs.")

    trials, metadata, weights, acclim = parse_all(paths)

    trials = add_metadata_context(trials, metadata)
    weights = add_metadata_context(weights, metadata)
    acclim = add_metadata_context(acclim, metadata)
    sessions = add_metadata_context(summarize_sessions(trials), metadata)
    animals = summarize_animals(trials, metadata)
    qc = build_qc_table(metadata, trials, weights)
    pca_scores, pca_loadings, pca_variance = pca_table(animals)

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    figures = outdir / "figures"
    figures.mkdir(exist_ok=True)
    save_tables(outdir, trials, sessions, animals, weights, acclim, qc, pca_scores, pca_loadings, pca_variance)
    write_excel_output(
        outdir / "four_choice_master_outputs.xlsx",
        trials,
        sessions,
        animals,
        weights,
        acclim,
        qc,
        pca_scores,
        pca_loadings,
        pca_variance,
    )
    write_readme(outdir, paths, trials, sessions, animals, weights, acclim, qc)

    if make_png:
        from make_png_figures import (
            plot_entries,
            plot_error_types as plot_error_types_png,
            plot_four_choice_reversal_arena,
            plot_pca as plot_pca_png,
            plot_reversal_choice_percent,
            plot_reversal_entries,
            plot_reversal_entry_percent,
            plot_reversal_error_types,
            plot_reversal_latency,
            plot_ttc as plot_ttc_png,
            plot_trials_to_criterion_mean_sem,
        )

        tables = outdir / "tables"
        plot_ttc_png(animals, figures)
        plot_four_choice_reversal_arena(tables, figures)
        plot_trials_to_criterion_mean_sem(animals, figures)
        plot_error_types_png(animals, figures)
        plot_reversal_error_types(animals, figures)
        plot_reversal_choice_percent(animals, figures)
        plot_entries(animals, figures)
        plot_reversal_latency(animals, figures)
        plot_reversal_entries(animals, figures)
        plot_reversal_entry_percent(animals, figures)
        plot_pca_png(tables, figures)

    return FourChoiceRunResult(
        outdir=outdir,
        input_workbooks=paths,
        trials=trials,
        metadata=metadata,
        weights=weights,
        acclim=acclim,
        sessions=sessions,
        animals=animals,
        qc=qc,
        pca_scores=pca_scores,
        pca_loadings=pca_loadings,
        pca_variance=pca_variance,
    )


def combine_animal_tables(csv_paths: list[str | Path], lab_column: str = "lab_source") -> pd.DataFrame:
    """Combine animal-level outputs across labs for later pooled analyses."""
    frames = []
    for path in csv_paths:
        path = Path(path)
        frame = pd.read_csv(path)
        if lab_column not in frame.columns:
            frame[lab_column] = path.parent.parent.name
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)
