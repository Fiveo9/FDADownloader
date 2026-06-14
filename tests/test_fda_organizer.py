import sys
import os
import tempfile
import types
import unittest


fake_pandas = types.ModuleType("pandas")
fake_pandas.isna = lambda value: value is None
sys.modules.setdefault("pandas", fake_pandas)

fake_tqdm = types.ModuleType("tqdm")
fake_tqdm.tqdm = lambda iterable, **kwargs: iterable
sys.modules.setdefault("tqdm", fake_tqdm)

from FDAOrganizer import FDAOrganizer


class FDAOrganizerCategoryTests(unittest.TestCase):
    def setUp(self):
        self.organizer = FDAOrganizer("dummy.xlsx")

    def test_category_uses_summary_topic_and_title_in_priority_order(self):
        row = {
            "FDA Organization": "Center for Devices and Radiological Health",
            "Topic": "Clinical Trials",
            "Summary": "Oncology drug combination development",
            "_Title_Internal": "Medical device guidance",
        }

        self.assertEqual("01_OCE_Oncology", self.organizer.determine_row_category(row))

    def test_loads_category_rules_from_csv_when_available(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            rules_path = os.path.join(temp_dir, "classification_rules.csv")
            with open(rules_path, "w", encoding="utf-8", newline="") as rules_file:
                rules_file.write("Keyword,Folder\nCUSTOM CENTER,10_Custom_Center\n")

            organizer = FDAOrganizer("dummy.xlsx", rules_path=rules_path)

        self.assertEqual("10_Custom_Center", organizer.determine_category("custom center policy"))


if __name__ == "__main__":
    unittest.main()
