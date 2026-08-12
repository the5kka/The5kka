import hashlib
import json
import re
import shutil
import sqlite3
import tempfile
import threading
import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path

import main_v4_1 as app


@contextmanager
def isolated_data(root: Path):
    names = [
        "DATA_DIR", "LEGACY_DATA_DIR", "LOG_DIR", "BACKUP_DIR", "EXPORT_DIR", "AUTO_BACKUP_DIR",
        "CONFIG_FILE", "CONFIG_BACKUP_FILE", "LOGO_FILE", "KCC_LOGO_FILE", "TLB_LOGO_FILE",
        "SIMTEK_LOGO_FILE", "LEGACY_CONFIG_FILE", "KCC_PKG_DATA_DIR", "KCC_PKG_DB_FILE",
        "TLB_DATA_DIR", "TLB_DB_FILE", "KCC_HDI_DATA_DIR", "KCC_HDI_DB_FILE",
        "SIMTEK_HDI_DATA_DIR", "SIMTEK_HDI_DB_FILE", "SIMTEK_HDI_CONDITION_MASTER_DB_FILE",
        "CONDITION_MASTER_DB_FILE", "LEGACY_KCC_PKG_DB_FILE", "LEGACY_CONDITION_MASTER_FILE",
        "MIGRATION_BACKUP_DONE", "MIGRATION_BACKUP_PATH",
    ]
    original = {name: getattr(app, name) for name in names}
    data = root / "data_v4_1"
    replacements = {
        "DATA_DIR": data,
        "LEGACY_DATA_DIR": root / "data",
        "LOG_DIR": data / "logs",
        "BACKUP_DIR": data / "backup",
        "EXPORT_DIR": data / "export",
        "AUTO_BACKUP_DIR": data / "auto_backup",
        "CONFIG_FILE": data / "config.json",
        "CONFIG_BACKUP_FILE": data / "config.json.bak",
        "LOGO_FILE": data / "company_logo.png",
        "KCC_LOGO_FILE": data / "korea_circuit_logo.png",
        "TLB_LOGO_FILE": data / "tlb_logo.png",
        "SIMTEK_LOGO_FILE": data / "simtek_logo.png",
        "LEGACY_CONFIG_FILE": root / "config.json",
        "KCC_PKG_DATA_DIR": data / "KCC_PKG",
        "KCC_PKG_DB_FILE": data / "KCC_PKG" / "work_log.db",
        "TLB_DATA_DIR": data / "TLB",
        "TLB_DB_FILE": data / "TLB" / "work_log.db",
        "KCC_HDI_DATA_DIR": data / "KCC_HDI",
        "KCC_HDI_DB_FILE": data / "KCC_HDI" / "work_log.db",
        "SIMTEK_HDI_DATA_DIR": data / "SIMTEK_HDI",
        "SIMTEK_HDI_DB_FILE": data / "SIMTEK_HDI" / "work_log.db",
        "SIMTEK_HDI_CONDITION_MASTER_DB_FILE": data / "SIMTEK_HDI" / "condition_master.db",
        "CONDITION_MASTER_DB_FILE": data / "KCC_PKG" / "condition_master.db",
        "LEGACY_KCC_PKG_DB_FILE": data / "KCC_PKG.db",
        "LEGACY_CONDITION_MASTER_FILE": root / "condition_master.json",
        "MIGRATION_BACKUP_DONE": False,
        "MIGRATION_BACKUP_PATH": None,
    }
    for name, value in replacements.items():
        setattr(app, name, value)
    try:
        yield data
    finally:
        for name, value in original.items():
            setattr(app, name, value)


def create_sample_db(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    try:
        conn.execute("CREATE TABLE sample (value TEXT)")
        conn.execute("INSERT INTO sample(value) VALUES (?)", (value,))
        conn.commit()
    finally:
        conn.close()


class V41RegressionTests(unittest.TestCase):
    def setUp(self):
        self._runtime_temp = tempfile.TemporaryDirectory()
        self._original_log_dir = app.LOG_DIR
        app.LOG_DIR = Path(self._runtime_temp.name) / "logs"

    def tearDown(self):
        app.LOG_DIR = self._original_log_dir
        self._runtime_temp.cleanup()

    def test_version_is_fully_separated(self):
        self.assertEqual(app.DATA_DIR.name, "data_v4_1")
        self.assertEqual(app.WINDOW_TITLE, "JIIN DNC Manager V4-1")
        self.assertIn("V4-1", app.APP_VERSION_TEXT)
        self.assertEqual(app.SINGLE_INSTANCE_MUTEX_NAME, "JIIN_DNC_Manager_V4_Single_Instance")
        self.assertEqual(app.LEGACY_SINGLE_INSTANCE_MUTEX_NAME, "JIIN_DNC_Manager_Single_Instance")

    def test_exact_search_and_atomic_dnc_copy(self):
        with tempfile.TemporaryDirectory() as temp, isolated_data(Path(temp)):
            root = Path(temp)
            source = root / "source"
            transfer = root / "transfer"
            nested = source / "nested"
            nested.mkdir(parents=True)
            transfer.mkdir()
            exact = nested / "KCC_001.txt"
            exact.write_text("G01 X1 Y2\n", encoding="utf-8")
            (nested / "KCC_001_REV1.txt").write_text("wrong", encoding="utf-8")
            (nested / "OLD_KCC_001.txt").write_text("wrong", encoding="utf-8")
            matches = app.search_condition_file("KCC_001", source)
            self.assertEqual(matches, [exact])
            copied = app.copy_dnc_file(exact, transfer)
            self.assertEqual(copied.read_bytes(), exact.read_bytes())
            self.assertEqual(hashlib.sha256(copied.read_bytes()).hexdigest(), hashlib.sha256(exact.read_bytes()).hexdigest())
            self.assertFalse(list(transfer.glob("*.jiin_part")))

    def test_missing_transfer_folder_is_never_created(self):
        with tempfile.TemporaryDirectory() as temp:
            missing = Path(temp) / "missing"
            with self.assertRaises(FileNotFoundError):
                app.delete_existing_dnc_txt(missing)
            self.assertFalse(missing.exists())

    def test_dnc_delete_thread_reports_verified_success(self):
        with tempfile.TemporaryDirectory() as temp:
            copied = Path(temp) / "active_dnc.txt"
            copied.write_text("DNC", encoding="utf-8")
            delete_thread = app.start_dnc_delete_thread(copied, 0)
            delete_thread.join(timeout=5)
            self.assertFalse(delete_thread.is_alive())
            self.assertFalse(copied.exists())
            self.assertTrue(app.dnc_delete_succeeded(delete_thread))

    def test_simtek_condition_parser_and_exact_file_selection(self):
        condition = "▷후가공(3차): 519X618MM ▷후가공GUIDE(3차):X110.380, Y604.000 (X -5.00MM SHIFT)"
        lot = {"round": "3차", "condition": condition, "additional_process": "Bond_2"}
        parsed, errors = app.parse_simtek_hdi_condition_source(lot)
        self.assertEqual(errors, [])
        self.assertEqual(parsed, {"size": "519x618", "shift": "-5", "jig": "604"})
        self.assertEqual(
            app.compose_simtek_hdi_condition_name(lot, parsed),
            "519x618_-5_5R_5up_One Zig_Bond_2",
        )
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            folder = root / "Trim Data Bond_2"
            folder.mkdir(parents=True)
            exact = folder / "519x618_-5_5R_5up_One Zig_Bond_2.txt"
            exact.write_text("DNC", encoding="utf-8")
            config = {"source_dnc_folders": {"심텍 HDI": str(root)}}
            trim, jig, file_errors = app.build_simtek_hdi_generated_condition(config, lot)
            self.assertEqual(trim, exact.stem)
            self.assertEqual(jig, "[통합지그] 604")
            self.assertEqual(file_errors, [])
            duplicate_dir = folder / "old"
            duplicate_dir.mkdir()
            shutil.copy2(exact, duplicate_dir / exact.name)
            _trim, _jig, file_errors = app.build_simtek_hdi_generated_condition(config, lot)
            self.assertTrue(any("2개 이상" in error for error in file_errors))

    def test_round_mismatch_and_two_lot_mixing_are_blocked(self):
        parsed, errors = app.parse_simtek_hdi_condition_source(
            {"round": "2차", "condition": "▷후가공(3차): 519x618mm ▷후가공guide(3차): X1, Y604"}
        )
        self.assertTrue(parsed)
        self.assertTrue(any("차수 불일치" in error for error in errors))
        lot1 = {"lot_no": "LOT-A", "condition": "KCC_1", "jig": "J1"}
        lot2 = {"lot_no": "LOT-A", "condition": "KCC_2", "jig": "J2"}
        joint_errors = app.validate_new_model_joint_lots([lot1, lot2])
        self.assertTrue(any("LOT No 동일" in error for error in joint_errors))
        self.assertTrue(any("작업조건 다름" in error for error in joint_errors))
        self.assertTrue(any("지그 다름" in error for error in joint_errors))

    def test_all_databases_are_included_in_migration_backup(self):
        with tempfile.TemporaryDirectory() as temp, isolated_data(Path(temp)):
            sources = {
                app.KCC_PKG_DB_FILE: "kcc log",
                app.CONDITION_MASTER_DB_FILE: "kcc master",
                app.TLB_DB_FILE: "tlb log",
                app.KCC_HDI_DB_FILE: "kcc hdi log",
                app.SIMTEK_HDI_DB_FILE: "simtek log",
                app.SIMTEK_HDI_CONDITION_MASTER_DB_FILE: "simtek master",
            }
            for path, value in sources.items():
                create_sample_db(path, value)
            backup = app.create_migration_backup_once()
            self.assertEqual(backup, app.create_migration_backup_once())
            relative_targets = [
                "KCC_PKG/work_log.db", "KCC_PKG/condition_master.db", "TLB/work_log.db",
                "KCC_HDI/work_log.db", "SIMTEK_HDI/work_log.db", "SIMTEK_HDI/condition_master.db",
            ]
            for relative in relative_targets:
                target = backup / relative
                self.assertTrue(target.exists(), relative)
                conn = sqlite3.connect(target)
                try:
                    self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")
                finally:
                    conn.close()

    def test_database_schema_indexes_and_quality_results(self):
        with tempfile.TemporaryDirectory() as temp, isolated_data(Path(temp)):
            conn = app.get_kcc_pkg_connection()
            try:
                indexes = {row[1] for row in conn.execute("PRAGMA index_list(dnc_logs)")}
                self.assertIn("idx_dnc_logs_master_sync", indexes)
            finally:
                conn.close()
            common = {
                "machine": "트리밍 1호기", "work_date": "2026-07-11", "shift_group": "C",
                "shift": "주간", "worker": "작업자",
            }
            lot = {
                "step": "100", "round": "1차", "manage_no": "M1", "lot_no": "L1", "qty": "0",
                "process_code": "T1", "condition": "KCC_TEST", "jig": "J1",
            }
            log_id = app.insert_new_model_db(common, lot, "조장")
            app.update_new_model_db(log_id, lot["condition"], False)
            conn = app.get_kcc_pkg_connection()
            try:
                row = conn.execute("SELECT status, condition_name, burr_result FROM dnc_logs WHERE id=?", (log_id,)).fetchone()
                self.assertEqual(row["status"], "완료")
                self.assertTrue(row["condition_name"].startswith("[검증 NG 발생]"))
                self.assertEqual(row["burr_result"], "Burr/초도품 이상")
            finally:
                conn.close()

    def test_config_and_excel_are_saved_atomically(self):
        with tempfile.TemporaryDirectory() as temp, isolated_data(Path(temp)):
            config = {"config_version": 1, "app_version": app.APP_VERSION_TEXT, "worker": "오국진"}
            app.save_config(config)
            self.assertEqual(json.loads(app.CONFIG_FILE.read_text(encoding="utf-8")), config)
            self.assertTrue(app.CONFIG_BACKUP_FILE.exists())
            self.assertFalse(list(app.DATA_DIR.glob("*.tmp")))

            workbook_path = Path(temp) / "atomic.xlsx"
            workbook = app.Workbook()
            workbook.active["A1"] = "old"
            workbook.save(workbook_path)
            workbook.close()
            workbook = app.load_workbook(workbook_path)
            workbook.active["A1"] = "new"
            app.save_workbook_safely(workbook, workbook_path)
            workbook.close()
            self.assertTrue(zipfile.is_zipfile(workbook_path))
            check = app.load_workbook(workbook_path, read_only=True, data_only=True)
            try:
                self.assertEqual(check.active["A1"].value, "new")
            finally:
                check.close()
            self.assertFalse(list(Path(temp).glob("*.jiin_tmp_*")))

    def test_actual_worklog_header_and_id_columns(self):
        path = Path.home() / "Desktop" / "타사 DNC 작업일보.xlsm"
        if not path.exists():
            self.skipTest("현장 작업일보 사본 없음")
        workbook = app.load_workbook(path, read_only=True, data_only=False, keep_vba=True, keep_links=False)
        try:
            self.assertEqual(app.get_log_header_row(workbook["KCC PKG"]), 6)
            self.assertEqual(app.get_log_header_row(workbook["TLB"]), 2)
            self.assertEqual(app.get_log_header_row(workbook["KCC HDI"]), 2)
            self.assertEqual(app.get_simtek_hdi_log_header_row(workbook["심텍 HDI"]), 2)
            self.assertEqual(workbook["TLB"].cell(2, app.EXCEL_EXPORT_ID_COLUMN).value, "DNC_LOG_ID")
            self.assertEqual(workbook["KCC HDI"].cell(2, app.EXCEL_EXPORT_ID_COLUMN).value, "DNC_LOG_ID")
            self.assertEqual(workbook["심텍 HDI"].cell(2, app.SIMTEK_HDI_EXPORT_ID_COLUMN).value, "DNC_LOG_ID")
            self.assertIn("h:mm", workbook["심텍 HDI"]["S4"].number_format.lower())
        finally:
            workbook.close()

    def test_actual_xlsm_roundtrip_copy_stays_openable(self):
        source = Path.home() / "Desktop" / "타사 DNC 작업일보.xlsm"
        if not source.exists():
            self.skipTest("현장 작업일보 사본 없음")
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / source.name
            shutil.copy2(source, target)
            with zipfile.ZipFile(source) as original_zip:
                original_names = set(original_zip.namelist())
            workbook = app.load_workbook(target, keep_vba=True, keep_links=False)
            app.save_workbook_safely(workbook, target)
            workbook.close()
            self.assertTrue(zipfile.is_zipfile(target))
            with zipfile.ZipFile(target) as saved_zip:
                saved_names = set(saved_zip.namelist())
            if "xl/vbaProject.bin" in original_names:
                self.assertIn("xl/vbaProject.bin", saved_names)
            check = app.load_workbook(target, read_only=True, data_only=True, keep_vba=True, keep_links=False)
            try:
                self.assertIn("KCC PKG", check.sheetnames)
                self.assertIn("심텍 HDI", check.sheetnames)
            finally:
                check.close()

    def test_background_workers_have_no_direct_after_zero_calls(self):
        source = Path(app.__file__).read_text(encoding="utf-8")
        self.assertIsNone(re.search(r"\.after\(\s*0\s*,", source))

    def test_ui_dispatch_queues_background_callback(self):
        original_thread_id = app.UI_THREAD_ID
        original_root = app.UI_ROOT
        while not app.UI_TASK_QUEUE.empty():
            app.UI_TASK_QUEUE.get_nowait()
            app.UI_TASK_QUEUE.task_done()
        called = []
        try:
            app.UI_THREAD_ID = threading.get_ident()

            def worker():
                app.run_on_ui(lambda: called.append("done"))

            thread = threading.Thread(target=worker)
            thread.start()
            thread.join()
            self.assertEqual(called, [])
            callback = app.UI_TASK_QUEUE.get_nowait()
            callback()
            app.UI_TASK_QUEUE.task_done()
            self.assertEqual(called, ["done"])
        finally:
            app.UI_THREAD_ID = original_thread_id
            app.UI_ROOT = original_root


if __name__ == "__main__":
    unittest.main(verbosity=2)
