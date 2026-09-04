"""Command-line entry point for the Four Choice TABS pipeline."""

from __future__ import annotations

import argparse
from pathlib import Path

from .pipeline import run_pipeline

from run_four_choice_analysis import print_qc_summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the Four Choice TABS analysis pipeline.")
    parser.add_argument("inputs", nargs="+", type=Path, help="Filled TABS workbook(s) or directories containing .xlsx files.")
    parser.add_argument("--outdir", type=Path, default=Path("outputs/four_choice_analysis"))
    parser.add_argument("--no-png", action="store_true", help="Skip PNG figure generation.")
    args = parser.parse_args()

    result = run_pipeline(args.inputs, args.outdir, make_png=not args.no_png)
    print(f"Parsed {len(result.input_workbooks)} workbook(s).")
    print(
        "Wrote "
        f"{len(result.trials)} trial rows, "
        f"{len(result.sessions)} session rows, "
        f"{len(result.animals)} animal rows, "
        f"{len(result.weights)} weight rows, and "
        f"{len(result.acclim)} acclimation rows."
    )
    print_qc_summary(result.qc)
    print(f"Output directory: {result.outdir.resolve()}")


if __name__ == "__main__":
    main()
