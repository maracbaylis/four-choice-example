# MAABCD Four Choice Analysis Example

This folder shows how to read filled Four Choice workbooks and convert them into
QC tables, trial/session/animal summary tables, figures, and PCA outputs.

The notebooks are the main entry point. Helper code in `src/` and `scripts/`
keeps the notebooks short and makes the calculations easier to check.

## Contents

- `notebooks/01_read_and_standardize_input.ipynb`: reads the complete example workbooks and inspects the parsed tables.
- `notebooks/02_quality_control_and_plots.ipynb`: reviews QC flags and writes figures from the complete example workbooks.
- `notebooks/03_pca_and_cross_lab_ready_outputs.ipynb`: shows the PCA output tables and the cross-lab merge pattern.
- `example_data/`: two anonymized complete example workbooks with discrimination, recall, and reversal rows.
- `tests/`: small checks for the core parsing and behavior rules.
- `src/` and `scripts/`: helper code used by the notebooks.
- `requirements.txt`: Python packages needed to run the example.
- `environment.yml`: optional conda environment.

## Setup

From this folder:

```bash
python3 -m pip install -r requirements.txt
```

Or with conda:

```bash
conda env create -f environment.yml
conda activate four-choice-tabs
```

Then open the notebooks in JupyterLab, Jupyter Notebook, VS Code, or another
notebook editor.

To run the small code checks:

```bash
python3 -m unittest discover -s tests
```

## What The Code Does

The parser reads user-entered values from the workbook and recomputes the
analysis outputs in Python. It does not rely on the workbook `SUMMARY` sheet or
Excel helper formulas for correctness, error type, trials to criterion, ZT, age,
percent baseline, or PCA.

The first ten parsed metadata fields are treated as primary grouping and sorting
variables, not behavior scores: `animal_id`, `strain`, `dob`,
`homecage_mice_at_testing`, `adversity_condition`, `genotype`,
`birth_litter_size_p05_20`, `homecage_rank`, `adversity_age_range`, and
`vendor`. These columns are kept on the left side of the output CSVs so large
multi-lab datasets can be grouped, filtered, or plotted consistently. Other
metadata fields are also retained when available, including newly added metadata
labels that are not yet part of the standard parser, so they can be reviewed
later as possible moderators or noise variables.

When the metadata or submission format needs review, the code writes
`tables/metadata_qc.csv` and prints a short QC summary with the animal ID, source
workbook, field, and issue.

## Biology And Data Assumptions

The current parser is written for the ARC Four Choice TABS workbook format. It
uses the `METADATA`, `WEIGHT`, `ACCLIM`, `DISC`, `RECALL`, and `REV`/`REVERSAL`
sheets. It reads raw user-entered trial values from trial number, bowl order,
odor chosen, latency, entry sequence, and notes columns, then recalculates the
derived behavior fields in Python.

For discrimination and recall, odor `A` is treated as the rewarded odor. For
reversal, odor `B` is treated as the newly rewarded odor, odor `A` as the
previously rewarded odor, odor `C` as the irrelevant odor, and odor `D` as the
novel odor. Reversal choices to `A` are counted as perseverative before the
first correct reversal trial and regressive after the first correct reversal
trial.

Entry sequences are interpreted as arena quadrant visits, with `1 = NW`,
`2 = NE`, `3 = SW`, and `4 = SE`. Those quadrants are then mapped back to odors
using the row-specific bowl order. Trials to criterion are calculated as the
first trial where the animal reaches 8 correct choices within the most recent
10 non-omission trials. Omissions are skipped when building the 10-trial
criterion window.

These assumptions have been checked against the current UCSC example workbooks.
If the public template changes, the most important items to re-check are the
quadrant numbering convention, the trials-to-criterion rule, and when regressive
errors begin.

## Outputs

Running the notebooks creates an `outputs/notebook_example/` folder with:

- `tables/trial_level.csv`
- `tables/session_level.csv`
- `tables/animal_level_outputs.csv`
- `tables/weight_long.csv`
- `tables/metadata_qc.csv`
- `tables/pca_scores.csv`
- `four_choice_master_outputs.xlsx`
- figures in `figures/`

The bundled GitHub example uses two anonymized complete workbooks with
discrimination, recall, and reversal behavior. Because the example includes more
than one complete animal, the PCA notebook writes PCA scores, loadings, and
variance tables. Treat that PCA as a format check and code example; biological
interpretation should use the larger multi-lab dataset.

For multi-lab analysis, run the same workflow once per lab and combine the
`animal_level_outputs.csv` files.
