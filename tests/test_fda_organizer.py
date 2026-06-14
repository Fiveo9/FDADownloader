import sys
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


if __name__ == "__main__":
    unittest.main()
