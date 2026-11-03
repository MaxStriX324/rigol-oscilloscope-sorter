"""Regression coverage for documented 4 analog + 16 digital CSV exports.

Synthetic fixture: validates our XLSX pathway, not a claim about a particular
physical Rigol BIN layout. Hardware export is separately user-confirmed.
"""
import csv
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import rigol_sort as sorter


class MultiChannelTests(unittest.TestCase):
    def test_signal_index_recognizes_analog_and_digital_headers(self):
        analog = [f"CH{i}V" for i in range(1, 5)]
        digital = [f"D{i}" for i in range(16)]
        headers = ["Time(s)", *analog, *digital, "Comments", "Marker"]
        self.assertEqual(sorter.signal_channels(headers), analog + digital)

    def test_twenty_channels_preserved_in_xlsx_and_archive_index(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            headers = ["Time(s)"] + [f"CH{i}V" for i in range(1, 5)] + [
                f"D{i}" for i in range(16)
            ]
            rows = [
                [0.0, 0.1, 0.2, 0.3, 0.4] + [i % 2 for i in range(16)],
                [0.001, 0.5, 0.6, 0.7, 0.8] + [(i + 1) % 2 for i in range(16)],
            ]
            incoming = root / "RigolDS10.csv"
            with incoming.open("w", newline="", encoding="utf-8") as stream:
                writer = csv.writer(stream)
                writer.writerow(headers)
                writer.writerows(rows)
            old = time.time() - 120
            os.utime(incoming, (old, old))
            with patch.object(sorter, "STABILITY_DELAY_SECONDS", 0):
                processed, failed = sorter.sort_once(root, min_age_seconds=60)
            self.assertEqual((processed, failed), (1, 0))
            xlsx_files = list(root.glob("20*/*.xlsx"))
            self.assertEqual(len(xlsx_files), 1)
            book = load_workbook(xlsx_files[0], read_only=True, data_only=True)
            try:
                sheet = book.active
                self.assertEqual(list(next(sheet.values)), headers)
                data = list(sheet.values)[1:]
                self.assertEqual(len(data), 2)
                self.assertEqual(len(data[0]), 21)
                self.assertEqual(data[0][5], 0)
                self.assertEqual(data[1][20], 0)
            finally:
                book.close()
            with (root / "_logs" / "sort_index.csv").open(
                newline="", encoding="utf-8-sig"
            ) as stream:
                index = list(csv.DictReader(stream, delimiter=";"))
            self.assertEqual(len(index), 1)
            self.assertEqual(index[0]["channels"].split(","), headers[1:])


if __name__ == "__main__":
    unittest.main()
