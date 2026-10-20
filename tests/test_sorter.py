import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import rigol_sort as sorter


class RigolSorterTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def create(self, name, contents=b'example', age_s=90):
        path = self.root / name
        path.write_bytes(contents)
        mod_time = time.time() - age_s
        os.utime(path, (mod_time, mod_time))
        return path

    def test_group_png_wave_setup_for_one_capture(self):
        self.create('RigolDS1.png')
        self.create('RigolDS1.csv', b'Time(s),CH1V\n0,0.15\n1,0.20\n')
        self.create('RigolDS1.stp', b'settings')
        groups = sorter.collect_groups(self.root)
        self.assertEqual(len(groups), 1)
        self.assertEqual(set(groups[0].files), {'.png', '.csv', '.stp'})

    def test_waits_full_minute_since_last_file_was_modified(self):
        self.create('RigolDS2.png', age_s=100)
        self.create('RigolDS2.stp', age_s=59)
        groups = sorter.collect_groups(self.root)
        with patch.object(sorter, 'STABILITY_DELAY_SECONDS', 0):
            self.assertEqual(sorter.ready_groups(groups, min_age_seconds=60), [])
        self.assertEqual(sorter.MIN_FILE_AGE_SECONDS, 60)

    def test_accepts_60_second_old_capture(self):
        self.create('RigolDS3.png', age_s=90)
        with patch.object(sorter, 'STABILITY_DELAY_SECONDS', 0):
            ready = sorter.ready_groups(sorter.collect_groups(self.root), min_age_seconds=60)
        self.assertEqual([g.stem for g in ready], ['RigolDS3'])

    def test_snapshot_group_sorted_with_shared_n(self):
        # Digital channels D0..D15 are part of the screenshot, not parsed by this sorter.
        self.create('RigolDS4.png', b'png_data')
        self.create('RigolDS4.stp', b'config')
        with patch.object(sorter, 'STABILITY_DELAY_SECONDS', 0):
            processed, failed = sorter.sort_once(self.root, min_age_seconds=60)
        self.assertEqual((processed, failed), (1, 0))
        day_folders = [p for p in self.root.iterdir() if p.is_dir() and p.name.startswith('20')]
        self.assertEqual(len(day_folders), 1)
        self.assertEqual(len(list(day_folders[0].glob('*.png'))), 1)
        self.assertEqual(len(list((self.root / 'Seting').rglob('*.stp'))), 1)
        self.assertEqual((self.root / '_logs' / 'rigol_counter.txt').read_text(), '1')

    def test_watch_mode_stops_cleanly_and_removes_lock(self):
        with patch('sys.argv', ['rigol_sort.py', '--root', str(self.root),
                                '--watch', '--min-age-seconds', '60',
                                '--poll-seconds', '5']):
            with patch.object(sorter.time, 'sleep', side_effect=KeyboardInterrupt):
                self.assertEqual(sorter.main(), 0)
        self.assertFalse((self.root / '_logs' / 'rigol_sort.lock').exists())

    def test_csv_translated_to_xlsx_and_original_kept(self):
        self.create('RigolDS5.csv', b'Time(s),CH1V\n0,0.15\n1,0.20\n')
        with patch.object(sorter, 'STABILITY_DELAY_SECONDS', 0):
            processed, failed = sorter.sort_once(self.root, min_age_seconds=60)
        self.assertEqual((processed, failed), (1, 0))
        self.assertEqual(len(list(self.root.glob('20*/*.csv'))), 1)
        self.assertEqual(len(list(self.root.glob('20*/*.xlsx'))), 1)


if __name__ == '__main__':
    unittest.main()
