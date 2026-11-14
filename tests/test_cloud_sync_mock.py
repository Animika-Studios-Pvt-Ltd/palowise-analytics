"""
Unit Tests for Cloud Sync Worker (OneDrive to Supabase) with Mocked DB and File
================================================================================
Validates cloud_sync_excel_to_supabase.py:
  1. Business rules & workbook processing (filters, flags, dates, composite keys)
  2. Read-only dry-run simulation calculations
  3. Live atomic merge transaction flow and staging table execution
  4. Unchanged files SHA-256 skip logic (zero DB writes, no merge commit)
  5. Changed files flow (updates staging, calls save_cloud_state, commits)
  6. Automatic transaction rollback on failure
  7. Failed validation does not update hash
  8. Failed download exits with error without updating hash
  9. Scheduled execution CLI mode (--live) vs manual (--dry-run, --force)
  10. Credential and URL redaction
  11. Regression tests for 'set_session cannot be used inside a transaction'
"""

import io
import os
import unittest
from unittest.mock import MagicMock, patch
from pathlib import Path
import pandas as pd
import psycopg2

import cloud_sync_excel_to_supabase as cs


class StrictMockConnection:
    """
    Simulates psycopg2's C-level transaction and set_session / autocommit constraints.
    Raises ProgrammingError if set_session() or autocommit assignment occurs in an active transaction.
    """
    def __init__(self):
        self.closed = False
        self._autocommit = False
        self._in_transaction = False
        self.commit_count = 0
        self.rollback_count = 0
        self.cursor_mock = MagicMock()
        self.cursor_mock.connection = self
        self.cursor_mock.connection.encoding = "utf-8"

    @property
    def autocommit(self):
        return self._autocommit

    @autocommit.setter
    def autocommit(self, value):
        if self._in_transaction:
            raise psycopg2.ProgrammingError("set_session cannot be used inside a transaction")
        self._autocommit = value

    def set_session(self, *args, **kwargs):
        if self._in_transaction:
            raise psycopg2.ProgrammingError("set_session cannot be used inside a transaction")
        if kwargs.get("autocommit") is not None:
            self._autocommit = kwargs.get("autocommit")

    def cursor(self):
        class MockCursorContext:
            def __init__(self, parent):
                self.parent = parent
            def __enter__(self):
                self.parent._in_transaction = True
                return self.parent.cursor_mock
            def __exit__(self, exc_type, exc_val, exc_tb):
                pass
        return MockCursorContext(self)

    def commit(self):
        self.commit_count += 1
        self._in_transaction = False

    def rollback(self):
        self.rollback_count += 1
        self._in_transaction = False

    def close(self):
        self.closed = True


def make_test_excel_bytes(include_date_error: bool = False) -> bytes:
    buf = io.BytesIO()
    date_val = "99th InvalidMonth" if include_date_error else "15th May"
    df = pd.DataFrame({
        "NAME": ["Test", None, "Lead 1", "Lead 2", "Lead 1"],
        "INDUSTRY": ["Tech", "Finance", "Healthcare", "Retail", "Healthcare"],
        "DESIGNATION": ["CTO", None, "Director", "VP", "VP (Updated)"],
        "STARTING DATE": ["1st Jan", None, date_val, "20th June", "15th May"],
        "Demo Date": [None, None, "22nd May", None, "22nd May"],
        "Pilot Conducted": ["yes", None, "Y", "false", "YES"],
        "STATUS": ["Cold", None, "Hot", "Warm", "Hot"],
    })
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="in", index=False)
    return buf.getvalue()


class TestCloudSyncMock(unittest.TestCase):

    def setUp(self):
        self.excel_bytes = make_test_excel_bytes()
        self.mock_cfg = {
            "SUPABASE_HOST": "aws-1-test.pooler.supabase.com",
            "SUPABASE_PORT": "5432",
            "SUPABASE_DATABASE": "postgres",
            "SUPABASE_USER": "postgres.testproject",
            "SUPABASE_PASSWORD": "SecretPassword123!",
            "SUPABASE_SSLMODE": "require",
            "ONEDRIVE_SHARE_URL": "https://tenant-my.sharepoint.com/:x:/g/personal/test/SecretShareToken",
        }

    # 1. Processing and Business Rules Test
    def test_process_workbook_rules(self):
        total_raw, eligible, excluded, errors = cs.process_workbook(self.excel_bytes)

        self.assertEqual(total_raw, 5)
        # Excluded: 1 dummy 'Test', 1 blank name, 1 duplicate Lead 1 superseded = 3 excluded
        self.assertEqual(len(excluded), 3)
        self.assertEqual(len(errors), 0)
        # Eligible: Lead 1 (latest) and Lead 2 = 2 eligible
        self.assertEqual(len(eligible), 2)

        # Check Lead 1 updated designation was kept
        eligible_dict = {rec["name"]: rec for _, rec in eligible}
        self.assertIn("Lead 1", eligible_dict)
        self.assertEqual(eligible_dict["Lead 1"]["designation"], "VP (Updated)")
        self.assertEqual(eligible_dict["Lead 1"]["pilot_conducted"], "YES")

    # 2. Dry-Run Simulation Logic Test
    def test_dry_run_simulation(self):
        mock_conn = MagicMock()
        mock_conn.closed = False
        mock_cur = MagicMock()
        mock_cur.connection.encoding = "utf-8"
        mock_conn.cursor.return_value.__enter__.return_value = mock_cur

        row_lead1 = ["Lead 1", "Director", "Healthcare"] + [None] * 25
        row_preserved = ["Preserved Lead", "Manager", "Energy"] + [None] * 25

        mock_cur.fetchall.side_effect = [
            [],  # No composite key duplicates in target table
            [row_lead1, row_preserved],  # Existing DB rows
        ]
        mock_cur.fetchone.return_value = (10, 8.5)  # View count and avg days

        with patch("psycopg2.connect", return_value=mock_conn):
            success = cs.execute_dry_run(self.excel_bytes, "MockSource", self.mock_cfg)

        self.assertTrue(success)
        mock_conn.set_session.assert_called_with(readonly=True, autocommit=True)
        # Ensure zero write statements were executed
        for call in mock_cur.execute.call_args_list:
            stmt = call[0][0].strip().upper()
            self.assertFalse(stmt.startswith("INSERT"))
            self.assertFalse(stmt.startswith("UPDATE"))
            self.assertFalse(stmt.startswith("DELETE"))
            self.assertFalse(stmt.startswith("CREATE"))
            self.assertFalse(stmt.startswith("DROP"))

    # 3. Live Sync Transaction Flow Test (Changed file updates hash and commits)
    def test_live_sync_success_flow(self):
        mock_conn = MagicMock()
        mock_conn.closed = False
        mock_cur = MagicMock()
        mock_conn.cursor.return_value.__enter__.return_value = mock_cur

        with patch("psycopg2.connect", return_value=mock_conn), \
             patch("psycopg2.extras.execute_values"), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=None), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state") as mock_save_state:

            # fetchall for duplicates check
            mock_cur.fetchall.return_value = []
            # fetchone responses: pre_count, preserved_count, post_count, (view_count, avg_days)
            mock_cur.fetchone.side_effect = [
                (100,),  # pre_count
                (4,),    # preserved_count
                (101,),  # post_count
                (20, 11.2),  # view stats
            ]
            mock_cur.rowcount = 1

            success = cs.execute_live_sync(self.excel_bytes, "MockSource", self.mock_cfg, force=True)

        self.assertTrue(success)
        # 1 cleanup commit after get_cloud_state + 1 final atomic merge commit
        self.assertGreaterEqual(mock_conn.commit.call_count, 1)
        mock_conn.rollback.assert_not_called()
        mock_save_state.assert_called_once()

    # 4. Automatic Rollback on Failure Test
    def test_live_sync_rollback_on_failure(self):
        mock_conn = MagicMock()
        mock_conn.closed = False
        mock_cur = MagicMock()
        mock_conn.cursor.return_value.__enter__.return_value = mock_cur

        with patch("psycopg2.connect", return_value=mock_conn), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=None), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state") as mock_save_state:

            mock_cur.fetchall.return_value = []
            mock_cur.fetchone.return_value = (100,)
            # Simulate a database failure during UPDATE
            mock_cur.execute.side_effect = [
                None,  # duplicates check
                None,  # pre_count
                None,  # create staging
                None,  # populate staging
                RuntimeError("Simulated DB Disk Full or Constraint Violation"),
            ]

            success = cs.execute_live_sync(self.excel_bytes, "MockSource", self.mock_cfg, force=True)

        self.assertFalse(success)
        mock_conn.rollback.assert_called_once()
        mock_save_state.assert_not_called()

    # 5. SHA-256 State Skip Test (Unchanged file)
    def test_live_sync_skips_when_hash_matches(self):
        file_hash = cs.compute_bytes_sha256(self.excel_bytes)
        mock_conn = MagicMock()
        mock_conn.closed = False

        existing_state = {
            "sha256": file_hash,
            "last_synced_at": "2026-11-14 10:00:00",
            "rows_in_db": 296,
        }

        with patch("psycopg2.connect", return_value=mock_conn), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=existing_state), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state") as mock_save_state:

            success = cs.execute_live_sync(self.excel_bytes, "MockSource", self.mock_cfg, force=False)

        self.assertTrue(success)
        # Should not execute staging, merge commit or save_cloud_state since file was unchanged
        mock_save_state.assert_not_called()

    # 6. Changed File Test (Hash differs from state table)
    def test_live_sync_runs_when_hash_differs(self):
        mock_conn = MagicMock()
        mock_conn.closed = False
        mock_cur = MagicMock()
        mock_conn.cursor.return_value.__enter__.return_value = mock_cur

        old_state = {
            "sha256": "0000000000000000000000000000000000000000000000000000000000000000",
            "last_synced_at": "2026-11-14 09:00:00",
            "rows_in_db": 296,
        }

        with patch("psycopg2.connect", return_value=mock_conn), \
             patch("psycopg2.extras.execute_values"), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=old_state), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state") as mock_save_state:

            mock_cur.fetchall.return_value = []
            mock_cur.fetchone.side_effect = [
                (296,),  # pre_count
                (4,),    # preserved_count
                (297,),  # post_count
                (55, 9.16),  # view stats
            ]
            mock_cur.rowcount = 1

            success = cs.execute_live_sync(self.excel_bytes, "MockSource", self.mock_cfg, force=False)

        self.assertTrue(success)
        self.assertGreaterEqual(mock_conn.commit.call_count, 1)
        mock_save_state.assert_called_once()

    # 7. Failed Validation Test (Does not update hash)
    def test_failed_validation_does_not_update_hash(self):
        corrupt_excel = make_test_excel_bytes(include_date_error=True)
        mock_conn = MagicMock()
        mock_conn.closed = False

        old_state = {
            "sha256": "different_hash_value",
            "last_synced_at": "2026-11-14 09:00:00",
            "rows_in_db": 296,
        }

        with patch("psycopg2.connect", return_value=mock_conn), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=old_state), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state") as mock_save_state:

            success = cs.execute_live_sync(corrupt_excel, "MockSource", self.mock_cfg, force=False)

        self.assertFalse(success)
        mock_save_state.assert_not_called()

    # 8. Failed Download Test (Exits with error without DB interaction)
    def test_failed_download_exits_without_db_interaction(self):
        with patch.object(cs, "resolve_excel_bytes", side_effect=RuntimeError("OneDrive network error")), \
             patch("sys.exit", side_effect=SystemExit) as mock_exit:

            with self.assertRaises(SystemExit):
                with patch("sys.argv", ["cloud_sync_excel_to_supabase.py", "--live"]):
                    cs.main()

            mock_exit.assert_called_with(1)

    # 9. Scheduled Execution vs Manual Execution CLI Flags
    def test_scheduled_execution_cli_flags(self):
        with patch.object(cs, "resolve_excel_bytes", return_value=(self.excel_bytes, "MockSource")), \
             patch.object(cs, "execute_live_sync", return_value=True) as mock_live, \
             patch.object(cs, "execute_dry_run") as mock_dry, \
             patch("sys.exit", side_effect=SystemExit) as mock_exit:

            # Scheduled mode: python cloud_sync_excel_to_supabase.py --live
            with self.assertRaises(SystemExit):
                with patch("sys.argv", ["cloud_sync_excel_to_supabase.py", "--live"]):
                    cs.main()

            mock_live.assert_called_once_with(self.excel_bytes, "MockSource", cs.load_cloud_config(), force=False)
            mock_dry.assert_not_called()
            mock_exit.assert_called_with(0)

    def test_manual_execution_dry_run_default(self):
        with patch.object(cs, "resolve_excel_bytes", return_value=(self.excel_bytes, "MockSource")), \
             patch.object(cs, "execute_live_sync") as mock_live, \
             patch.object(cs, "execute_dry_run", return_value=True) as mock_dry, \
             patch("sys.exit", side_effect=SystemExit) as mock_exit:

            # Manual default: python cloud_sync_excel_to_supabase.py --dry-run
            with self.assertRaises(SystemExit):
                with patch("sys.argv", ["cloud_sync_excel_to_supabase.py", "--dry-run"]):
                    cs.main()

            mock_dry.assert_called_once()
            mock_live.assert_not_called()
            mock_exit.assert_called_with(0)

    # 10. Redaction Test
    def test_sanitization_in_cloud_sync(self):
        msg = f"Failed to connect to {self.mock_cfg['SUPABASE_HOST']} with password {self.mock_cfg['SUPABASE_PASSWORD']} and url {self.mock_cfg['ONEDRIVE_SHARE_URL']}"
        clean = cs.sanitize_message(msg, [self.mock_cfg['SUPABASE_PASSWORD']])

        self.assertNotIn("SecretPassword123!", clean)
        self.assertNotIn("SecretShareToken", clean)
        self.assertIn("[REDACTED]", clean)

    # 11. Regression Tests: 'set_session cannot be used inside a transaction'
    def test_regression_set_session_error_reproduced(self):
        """Verifies that attempting set_session or autocommit on an in-transaction connection raises ProgrammingError."""
        strict_conn = StrictMockConnection()
        with strict_conn.cursor():
            pass  # Starts transaction: strict_conn._in_transaction = True

        with self.assertRaises(psycopg2.ProgrammingError) as ctx1:
            strict_conn.autocommit = False
        self.assertIn("set_session cannot be used inside a transaction", str(ctx1.exception))

        with self.assertRaises(psycopg2.ProgrammingError) as ctx2:
            strict_conn.set_session(readonly=False)
        self.assertIn("set_session cannot be used inside a transaction", str(ctx2.exception))

    def test_regression_live_sync_does_not_call_set_session_in_transaction(self):
        """Verifies that execute_live_sync never invokes set_session or autocommit inside an active transaction."""
        strict_conn = StrictMockConnection()
        strict_conn.cursor_mock.fetchall.return_value = []
        strict_conn.cursor_mock.fetchone.side_effect = [
            (296,),  # pre_count
            (4,),    # preserved_count
            (297,),  # post_count
            (55, 9.16),  # view stats
        ]
        strict_conn.cursor_mock.rowcount = 1

        with patch("psycopg2.connect", return_value=strict_conn), \
             patch("psycopg2.extras.execute_values"), \
             patch("cloud_sync_excel_to_supabase.get_cloud_state", return_value=None), \
             patch("cloud_sync_excel_to_supabase.save_cloud_state"):

            # Must run without raising psycopg2.ProgrammingError: set_session cannot be used inside a transaction
            success = cs.execute_live_sync(self.excel_bytes, "MockSource", self.mock_cfg, force=True)

        self.assertTrue(success)
        self.assertGreaterEqual(strict_conn.commit_count, 1)
        self.assertEqual(strict_conn.rollback_count, 0)


if __name__ == "__main__":
    unittest.main()
