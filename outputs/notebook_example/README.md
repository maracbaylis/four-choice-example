# Four Choice TABS Analysis Run

Input workbooks: 2
Trial rows: 94
Session rows: 6
Animal rows: 2
Weight rows: 14
Acclimation rows: 6
Metadata QC flags: 0

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

Some figure types require recall or reversal rows. If those sheets are empty in
the input workbooks, the corresponding figures are skipped.
