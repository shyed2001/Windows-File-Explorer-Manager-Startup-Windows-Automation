"""
test_file_explorer_manager.py
-----------------------------
Unit tests for file_explorer_manager.py

These tests run on any OS (including Linux CI) by mocking all Windows-specific
calls so that the core logic can be exercised in a portable way.

Run with:
  python -m pytest test_file_explorer_manager.py -v
"""

import json
import os
import sys
import tempfile
import types
import unittest
from unittest.mock import MagicMock, patch, call


# ---------------------------------------------------------------------------
# Provide a minimal ctypes.windll stub so the module can be imported on non-
# Windows environments (Linux CI runners, macOS, etc.).
# ---------------------------------------------------------------------------

if not hasattr(sys.modules.get("ctypes", None) or __import__("ctypes"), "windll"):
    import ctypes as _ctypes
    _ctypes.windll = types.SimpleNamespace(
        user32=MagicMock()
    )

import file_explorer_manager as fem  # noqa: E402  (imported after stub)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_config(windows: list, settings: dict | None = None) -> str:
    """Write a temporary config.json and return its path."""
    data = {"windows": windows}
    if settings:
        data["settings"] = settings
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    )
    json.dump(data, tmp)
    tmp.close()
    return tmp.name


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestExpandPath(unittest.TestCase):
    def test_tilde_expansion(self):
        result = fem._expand("~/Documents")
        self.assertNotIn("~", result)

    def test_env_var_expansion(self):
        os.environ["TEST_FEM_VAR"] = "hello"
        result = fem._expand("%TEST_FEM_VAR%\\world")
        # On non-Windows os.path.expandvars uses $VAR syntax, not %VAR%,
        # so on non-Windows this stays unexpanded – just check no exception.
        self.assertIsInstance(result, str)

    def test_literal_path_unchanged(self):
        path = "/some/absolute/path"
        self.assertEqual(fem._expand(path), path)


class TestOpenExplorerWindowDryRun(unittest.TestCase):
    """Dry-run mode must never call subprocess or ctypes."""

    def test_dry_run_prints_path(self):
        with patch("builtins.print") as mock_print, \
             patch("subprocess.Popen") as mock_popen:
            fem.open_explorer_window(
                "/tmp/nonexistent", position=None, dry_run=True
            )
            mock_popen.assert_not_called()
            printed = " ".join(str(a) for c in mock_print.call_args_list for a in c[0])
            self.assertIn("DRY-RUN", printed)

    def test_dry_run_includes_position(self):
        pos = {"left": 10, "top": 20, "width": 800, "height": 600}
        with patch("builtins.print") as mock_print, \
             patch("subprocess.Popen"):
            fem.open_explorer_window("/tmp", position=pos, dry_run=True)
            printed = " ".join(str(a) for c in mock_print.call_args_list for a in c[0])
            self.assertIn("position", printed)


class TestOpenExplorerWindowSkipMissing(unittest.TestCase):
    """Non-existent paths must be skipped gracefully."""

    def test_skips_nonexistent_path(self):
        with patch("builtins.print") as mock_print, \
             patch("subprocess.Popen") as mock_popen:
            fem.open_explorer_window("/path/that/does/not/exist/xyz123")
            mock_popen.assert_not_called()
            printed = " ".join(str(a) for c in mock_print.call_args_list for a in c[0])
            self.assertIn("WARN", printed)


class TestRunFromConfigDryRun(unittest.TestCase):
    """run_from_config with dry_run=True should open nothing."""

    def _config_with_tmpdir(self):
        """Return a config whose path actually exists (tmpdir)."""
        return _make_config([
            {"path": tempfile.gettempdir(), "label": "Temp"},
        ])

    def test_dry_run_no_popen(self):
        cfg = self._config_with_tmpdir()
        try:
            with patch("subprocess.Popen") as mock_popen, \
                 patch("builtins.print"):
                fem.run_from_config(cfg, dry_run=True)
            mock_popen.assert_not_called()
        finally:
            os.unlink(cfg)

    def test_dry_run_prints_dry_run_message(self):
        cfg = self._config_with_tmpdir()
        try:
            with patch("builtins.print") as mock_print, \
                 patch("subprocess.Popen"):
                fem.run_from_config(cfg, dry_run=True)
            printed = " ".join(str(a) for c in mock_print.call_args_list for a in c[0])
            self.assertIn("DRY-RUN", printed)
        finally:
            os.unlink(cfg)


class TestRunFromConfigMissingFile(unittest.TestCase):
    def test_exits_on_missing_config(self):
        with self.assertRaises(SystemExit):
            fem.run_from_config("/no/such/config_xyz.json")


class TestRunFromConfigNoWindows(unittest.TestCase):
    def test_empty_windows_list(self):
        cfg = _make_config([])
        try:
            with patch("builtins.print") as mock_print, \
                 patch("subprocess.Popen") as mock_popen:
                fem.run_from_config(cfg)
            mock_popen.assert_not_called()
            printed = " ".join(str(a) for c in mock_print.call_args_list for a in c[0])
            self.assertIn("WARN", printed)
        finally:
            os.unlink(cfg)


class TestRunFromConfigDelay(unittest.TestCase):
    """Verify delay is applied between windows."""

    def test_delay_called_between_windows(self):
        tmpdir = tempfile.gettempdir()
        cfg = _make_config(
            [
                {"path": tmpdir, "label": "A"},
                {"path": tmpdir, "label": "B"},
            ],
            settings={"delay_between_windows_ms": 100},
        )
        try:
            with patch("subprocess.Popen"), \
                 patch("time.sleep") as mock_sleep, \
                 patch("time.monotonic", return_value=0.0), \
                 patch("builtins.print"):
                fem.run_from_config(cfg, dry_run=False)
            # sleep should have been called once (between 2 windows)
            mock_sleep.assert_called_once_with(0.1)
        finally:
            os.unlink(cfg)


class TestCLIParser(unittest.TestCase):
    def test_default_config_path(self):
        parser = fem._build_parser()
        args = parser.parse_args([])
        self.assertTrue(args.config.endswith("config.json"))

    def test_custom_config_path(self):
        parser = fem._build_parser()
        args = parser.parse_args(["--config", "/tmp/my.json"])
        self.assertEqual(args.config, "/tmp/my.json")

    def test_dry_run_flag(self):
        parser = fem._build_parser()
        args = parser.parse_args(["--dry-run"])
        self.assertTrue(args.dry_run)

    def test_no_dry_run_by_default(self):
        parser = fem._build_parser()
        args = parser.parse_args([])
        self.assertFalse(args.dry_run)


class TestMainEntryPoint(unittest.TestCase):
    def test_main_returns_zero_on_success(self):
        cfg = _make_config([{"path": tempfile.gettempdir(), "label": "T"}])
        try:
            with patch("subprocess.Popen"), \
                 patch("time.sleep"), \
                 patch("time.monotonic", return_value=0.0), \
                 patch("builtins.print"):
                rc = fem.main(["--config", cfg, "--dry-run"])
            self.assertEqual(rc, 0)
        finally:
            os.unlink(cfg)


if __name__ == "__main__":
    unittest.main()
