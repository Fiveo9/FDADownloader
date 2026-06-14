import sys
import os
import tempfile
import types
import unittest


fake_pandas = types.ModuleType("pandas")
fake_pandas.isna = lambda value: value is None
fake_pandas.notna = lambda value: value is not None
sys.modules.setdefault("pandas", fake_pandas)

fake_tqdm = types.ModuleType("tqdm")


class FakeTqdm:
    def __init__(self, iterable):
        self.iterable = iterable

    def __iter__(self):
        return iter(self.iterable)

    def set_description(self, description):
        self.description = description


fake_tqdm.tqdm = lambda iterable, **kwargs: FakeTqdm(iterable)
sys.modules.setdefault("tqdm", fake_tqdm)

from FDAOrganizer import FDAOrganizer, parse_args


class Row(dict):
    def to_dict(self):
        return dict(self)


class FakeDataFrame:
    def __init__(self, rows):
        self.rows = rows
        self.shape = (len(rows), len(rows[0]) if rows else 0)

    def iterrows(self):
        for index, row in enumerate(self.rows):
            yield index, row


class FDAOrganizerCategoryTests(unittest.TestCase):
    def setUp(self):
        self.organizer = FDAOrganizer("dummy.xlsx")

    def test_parse_args_accepts_paths_and_rules(self):
        args = parse_args([
            "--excel", "FDA_Guidance_Data_20260529.xlsx",
            "--source", "FDA_Downloads",
            "--target", "FDA_Guidance_Library",
            "--rules", "classification_rules.csv",
            "--dry-run",
        ])

        self.assertEqual("FDA_Guidance_Data_20260529.xlsx", args.excel)
        self.assertEqual("FDA_Downloads", args.source)
        self.assertEqual("FDA_Guidance_Library", args.target)
        self.assertEqual("classification_rules.csv", args.rules)
        self.assertTrue(args.dry_run)

    def test_category_uses_summary_topic_and_title_in_priority_order(self):
        row = {
            "FDA Organization": "Center for Devices and Radiological Health",
            "Topic": "Clinical Trials",
            "Summary": "Oncology drug combination development",
            "_Title_Internal": "Medical device guidance",
        }

        self.assertEqual("01_OCE_Oncology", self.organizer.determine_row_category(row))

    def test_reconstruct_filename_removes_invalid_path_characters(self):
        row = {
            "Summary": 'Guidance: With / Bad? "Characters"\n',
            "Issue Date": None,
        }

        filename = self.organizer.reconstruct_filename(row)

        self.assertTrue(filename.startswith("00000000_"))
        for invalid_char in '\\/*?:"<>|\r\n\t':
            self.assertNotIn(invalid_char, filename)

    def test_find_matching_file_accepts_truncated_download_name(self):
        row = {
            "Summary": "Clinical Pharmacology Considerations for Oligonucleotide Therapeutics",
            "Issue Date": None,
        }
        available_files = [
            "00000000_Clinical Pharmacology Considerations.pdf",
            "00000000_Unrelated Guidance.pdf",
        ]

        self.assertEqual(
            "00000000_Clinical Pharmacology Considerations.pdf",
            self.organizer.find_matching_file(row, available_files),
        )

    def test_loads_category_rules_from_csv_when_available(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            rules_path = os.path.join(temp_dir, "classification_rules.csv")
            with open(rules_path, "w", encoding="utf-8", newline="") as rules_file:
                rules_file.write("Keyword,Folder\nCUSTOM CENTER,10_Custom_Center\n")

            organizer = FDAOrganizer("dummy.xlsx", rules_path=rules_path)

        self.assertEqual("10_Custom_Center", organizer.determine_category("custom center policy"))

    def test_dry_run_previews_archive_without_copying_or_writing_index(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = os.path.join(temp_dir, "source")
            target_dir = os.path.join(temp_dir, "target")
            os.makedirs(source_dir)

            source_file = os.path.join(source_dir, "00000000_Dry Run Guidance.pdf")
            with open(source_file, "w", encoding="utf-8") as f:
                f.write("placeholder")

            row = Row({
                "Summary": "Dry Run Guidance",
                "Topic": "",
                "Issue Date": None,
                "FDA Organization": "",
            })

            original_read_excel = getattr(fake_pandas, "read_excel", None)
            fake_pandas.read_excel = lambda path: FakeDataFrame([row])
            try:
                organizer = FDAOrganizer(
                    "dummy.xlsx",
                    source_dir=source_dir,
                    target_dir=target_dir,
                    dry_run=True,
                )
                organizer.run()
            finally:
                if original_read_excel is None:
                    delattr(fake_pandas, "read_excel")
                else:
                    fake_pandas.read_excel = original_read_excel

            preview_path = os.path.join(target_dir, "Uncategorized", "00000000_Dry Run Guidance.pdf")
            index_path = os.path.join(target_dir, "00_FDA_Guidance_Index.xlsx")

            self.assertFalse(os.path.exists(preview_path))
            self.assertFalse(os.path.exists(index_path))
            self.assertEqual("预览归档", organizer.sheet_data["Uncategorized"][0]["Local Status"])


if __name__ == "__main__":
    unittest.main()
