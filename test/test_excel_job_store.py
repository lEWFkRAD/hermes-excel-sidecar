import importlib.util
import pathlib
import tempfile
import time
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("excel_job_store", ROOT / "jobs" / "excel_job_store.py")
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)
ExcelJobStore = module.ExcelJobStore


class ExcelJobStoreTests(unittest.TestCase):
    def test_create_checkpoint_and_get(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ExcelJobStore(pathlib.Path(directory) / "jobs.db")
            job = store.create(workbook_id="wb", sheet="Input", input_range="D10:D4009", output_ranges=["X10:Z4009"], cursor=10, batch_size=25, release="r1", session_id="s1", job_id="job_test")
            self.assertEqual(job["status"], "queued")
            lease = store.lease("job_test", seconds=10)
            updated = store.checkpoint("job_test", cursor=35, lease_id=lease)
            self.assertEqual(updated["cursor"], 35)
            self.assertTrue(updated["last_checkpoint"])

    def test_single_active_lease_and_recovery(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ExcelJobStore(pathlib.Path(directory) / "jobs.db")
            store.create(workbook_id="wb", sheet="Input", input_range="D1:D2", output_ranges=["X1:Z2"], cursor=1, batch_size=2, release="r1", session_id="s1", job_id="job_test")
            store.lease("job_test", seconds=1)
            with self.assertRaises(ValueError):
                store.lease("job_test", seconds=1)
            time.sleep(1.05)
            self.assertEqual(store.recover_stale(), 1)
            self.assertEqual(store.get("job_test")["status"], "paused")


if __name__ == "__main__":
    unittest.main()
