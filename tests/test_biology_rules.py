from __future__ import annotations

from pathlib import Path
import sys
import unittest

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

from four_choice_core import (  # noqa: E402
    add_sequence_metrics,
    metadata_qc,
    parse_entry_sequence,
    reversal_error_type,
    trials_to_criterion,
)


class BiologyRuleTests(unittest.TestCase):
    def test_entry_sequence_digits_map_to_quadrants_and_odors(self) -> None:
        quadrants, odors = parse_entry_sequence("1234", "ABCD")

        self.assertEqual(quadrants, ["NW", "NE", "SW", "SE"])
        self.assertEqual(odors, ["A", "B", "C", "D"])

    def test_entry_sequence_uses_trial_specific_bowl_order(self) -> None:
        quadrants, odors = parse_entry_sequence("1234", "BCDA")

        self.assertEqual(quadrants, ["NW", "NE", "SW", "SE"])
        self.assertEqual(odors, ["B", "C", "D", "A"])

    def test_ttc_is_8_correct_out_of_10_non_omission_trials(self) -> None:
        trials = pd.DataFrame(
            {
                "animal_id": ["animal_1"] * 11,
                "phase": ["discrimination"] * 11,
                "trial": list(range(1, 12)),
                "correct": [True, True, False, True, False, True, True, False, True, True, True],
                "omission": [False, False, False, False, True, False, False, False, False, False, False],
            }
        )

        scored = add_sequence_metrics(trials)

        self.assertEqual(scored.loc[scored["trial"] == 11, "correct_in_last_10_non_omission"].item(), 8)
        self.assertEqual(trials_to_criterion(scored), 11)

    def test_reversal_a_errors_switch_from_perseverative_to_regressive(self) -> None:
        self.assertEqual(reversal_error_type("A", previous_rewarded="A", new_rewarded="B", seen_correct=False), "Perseverative")
        self.assertEqual(reversal_error_type("A", previous_rewarded="A", new_rewarded="B", seen_correct=True), "Regressive")
        self.assertEqual(reversal_error_type("C", previous_rewarded="A", new_rewarded="B", seen_correct=True), "Irrelevant")
        self.assertEqual(reversal_error_type("D", previous_rewarded="A", new_rewarded="B", seen_correct=True), "Novel")

    def test_metadata_qc_flags_missing_required_fields(self) -> None:
        metadata = pd.DataFrame(
            [
                {
                    "animal_id": "animal_1",
                    "experiment_cohort": "ARC",
                    "adversity_condition": "N/A",
                    "sex": "",
                    "genotype": "Wildtype",
                    "lab_source": "Example Lab",
                    "animal_id_inferred_from_filename": False,
                }
            ]
        )

        qc = metadata_qc(metadata)

        self.assertIn("missing_required_metadata", set(qc["qc_flag"]))
        self.assertIn("sex", set(qc["field"]))


if __name__ == "__main__":
    unittest.main()
