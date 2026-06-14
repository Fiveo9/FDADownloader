import csv
import os
import sys
import tempfile
import types
import unittest


fake_pandas = types.ModuleType("pandas")
fake_pandas.notna = lambda value: value is not None
sys.modules.setdefault("pandas", fake_pandas)

fake_bs4 = types.ModuleType("bs4")
fake_bs4.BeautifulSoup = object
sys.modules.setdefault("bs4", fake_bs4)

fake_tqdm = types.ModuleType("tqdm")
fake_tqdm.tqdm = lambda iterable, **kwargs: iterable
sys.modules.setdefault("tqdm", fake_tqdm)

fake_openpyxl = types.ModuleType("openpyxl")
fake_openpyxl.utils = types.SimpleNamespace(get_column_letter=lambda index: "A")
fake_openpyxl_styles = types.ModuleType("openpyxl.styles")
fake_openpyxl_styles.Alignment = object
sys.modules.setdefault("openpyxl", fake_openpyxl)
sys.modules.setdefault("openpyxl.styles", fake_openpyxl_styles)

fake_selenium = types.ModuleType("selenium")
fake_webdriver = types.ModuleType("selenium.webdriver")
fake_chrome = types.ModuleType("selenium.webdriver.chrome")
fake_service = types.ModuleType("selenium.webdriver.chrome.service")
fake_options = types.ModuleType("selenium.webdriver.chrome.options")
fake_service.Service = object
fake_options.Options = object
fake_webdriver.chrome = fake_chrome
fake_chrome.service = fake_service
fake_chrome.options = fake_options
fake_selenium.webdriver = fake_webdriver
sys.modules.setdefault("selenium", fake_selenium)
sys.modules.setdefault("selenium.webdriver", fake_webdriver)
sys.modules.setdefault("selenium.webdriver.chrome", fake_chrome)
sys.modules.setdefault("selenium.webdriver.chrome.service", fake_service)
sys.modules.setdefault("selenium.webdriver.chrome.options", fake_options)

fake_webdriver_manager = types.ModuleType("webdriver_manager")
fake_webdriver_manager_chrome = types.ModuleType("webdriver_manager.chrome")
fake_webdriver_manager_chrome.ChromeDriverManager = object
sys.modules.setdefault("webdriver_manager", fake_webdriver_manager)
sys.modules.setdefault("webdriver_manager.chrome", fake_webdriver_manager_chrome)

from FDADownloader import FDADownloader


class FDADownloaderFailureLogTests(unittest.TestCase):
    def test_download_manifest_writes_status_rows_to_csv(self):
        with tempfile.TemporaryDirectory() as download_dir:
            downloader = FDADownloader("https://example.test", download_dir=download_dir)
            item = {
                "Summary": "Already downloaded guidance",
                "Topic": "Drugs",
                "Issue Date": "2026-02-03",
                "Download URL": "https://example.test/already.pdf",
            }

            downloader.record_download_status(
                item,
                status="skipped_existing",
                local_path=os.path.join(download_dir, "20260203_Already downloaded guidance.pdf"),
            )
            manifest_path = os.path.join(download_dir, "manifest.csv")

            self.assertEqual(manifest_path, downloader.save_download_manifest(manifest_path))
            with open(manifest_path, newline="", encoding="utf-8-sig") as manifest_file:
                rows = list(csv.DictReader(manifest_file))

        self.assertEqual(1, len(rows))
        self.assertEqual("2026-02-03", rows[0]["Issue Date"])
        self.assertEqual("Already downloaded guidance", rows[0]["Summary"])
        self.assertEqual("https://example.test/already.pdf", rows[0]["Download URL"])
        self.assertEqual("skipped_existing", rows[0]["Status"])
        self.assertTrue(rows[0]["Local Path"].endswith("20260203_Already downloaded guidance.pdf"))
        self.assertEqual("", rows[0]["Reason"])

    def test_failure_report_writes_download_context_to_csv(self):
        with tempfile.TemporaryDirectory() as download_dir:
            downloader = FDADownloader("https://example.test", download_dir=download_dir)
            item = {
                "Summary": "A guidance with / unsafe filename characters",
                "Topic": "Clinical",
                "Issue Date": "2026-01-02",
                "Download URL": "https://example.test/file.pdf",
            }

            downloader.record_download_failure(item, "download timeout")
            report_path = os.path.join(download_dir, "failures.csv")

            self.assertEqual(report_path, downloader.save_download_failures(report_path))
            with open(report_path, newline="", encoding="utf-8-sig") as report_file:
                rows = list(csv.DictReader(report_file))

        self.assertEqual(1, len(rows))
        self.assertEqual("2026-01-02", rows[0]["Issue Date"])
        self.assertEqual("A guidance with / unsafe filename characters", rows[0]["Summary"])
        self.assertEqual("Clinical", rows[0]["Topic"])
        self.assertEqual("https://example.test/file.pdf", rows[0]["Download URL"])
        self.assertEqual("download timeout", rows[0]["Reason"])


if __name__ == "__main__":
    unittest.main()
