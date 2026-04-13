"""
file_explorer_manager.py
------------------------
Windows File Explorer Manager – Startup Automation
Opens one or more File Explorer windows at configured paths and positions them
on screen automatically at Windows startup (or on demand).

Requirements:
  Python 3.8+  (standard library only)
  Windows OS   (uses win32 APIs via ctypes / subprocess)

Usage:
  python file_explorer_manager.py [--config config.json] [--dry-run]

Author: shyed2001
"""

import argparse
import ctypes
import json
import os
import subprocess
import sys
import time


# ---------------------------------------------------------------------------
# Constants & Win32 helpers
# ---------------------------------------------------------------------------

SW_RESTORE = 9
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010

# SetWindowPos flags
SWP_MOVE_RESIZE = SWP_NOZORDER | SWP_NOACTIVATE


def _expand(path: str) -> str:
    """Expand environment variables and user tilde in a path string."""
    return os.path.expandvars(os.path.expanduser(path))


def _set_window_pos(hwnd: int, left: int, top: int, width: int, height: int) -> bool:
    """Move and resize a window identified by its handle using SetWindowPos."""
    user32 = ctypes.windll.user32
    result = user32.SetWindowPos(
        hwnd,
        0,          # hWndInsertAfter (ignored when SWP_NOZORDER is set)
        left, top, width, height,
        SWP_MOVE_RESIZE,
    )
    return bool(result)


def _find_explorer_hwnd(target_path: str, timeout: float = 5.0) -> int:
    """
    Poll for a File Explorer window whose title contains the last component of
    *target_path*.  Returns the HWND on success, 0 on timeout.
    """
    user32 = ctypes.windll.user32
    folder_name = os.path.basename(target_path.rstrip("\\/")) or target_path

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        # Walk all top-level windows via GetWindow
        hwnd = user32.GetTopWindow(0)
        while hwnd:
            buf = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, buf, 512)
            if folder_name.lower() in buf.value.lower():
                return hwnd
            hwnd = user32.GetWindow(hwnd, 2)  # GW_HWNDNEXT = 2
        time.sleep(0.2)
    return 0


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def open_explorer_window(path: str, position: dict | None = None, dry_run: bool = False) -> None:
    """
    Open a single File Explorer window at *path* and optionally position it.

    :param path:     Absolute path to open (environment variables are expanded).
    :param position: Optional dict with keys: left, top, width, height (pixels).
    :param dry_run:  When True, print the action without executing it.
    """
    expanded = _expand(path)

    if dry_run:
        pos_info = f" -> position {position}" if position else ""
        print(f"[DRY-RUN] Would open: {expanded}{pos_info}")
        return

    if not os.path.isdir(expanded):
        print(f"[WARN] Path does not exist, skipping: {expanded}")
        return

    print(f"[INFO] Opening: {expanded}")
    subprocess.Popen(["explorer.exe", expanded])

    if position:
        hwnd = _find_explorer_hwnd(expanded)
        if hwnd:
            _set_window_pos(
                hwnd,
                position.get("left", 0),
                position.get("top", 0),
                position.get("width", 800),
                position.get("height", 600),
            )
        else:
            print(f"[WARN] Could not find window handle for: {expanded}")


def run_from_config(config_path: str, dry_run: bool = False) -> None:
    """
    Read *config_path* and open all configured File Explorer windows.

    :param config_path: Path to the JSON configuration file.
    :param dry_run:     When True, print actions without executing them.
    """
    config_path = os.path.abspath(config_path)
    if not os.path.isfile(config_path):
        print(f"[ERROR] Config file not found: {config_path}")
        sys.exit(1)

    with open(config_path, encoding="utf-8") as fh:
        config = json.load(fh)

    settings = config.get("settings", {})
    delay_ms = settings.get("delay_between_windows_ms", 300)

    windows = config.get("windows", [])
    if not windows:
        print("[WARN] No windows defined in config.")
        return

    for idx, win_cfg in enumerate(windows):
        path = win_cfg.get("path", "")
        label = win_cfg.get("label", path)
        position = win_cfg.get("position")  # may be None

        if not path:
            print(f"[WARN] Window #{idx} has no path defined, skipping.")
            continue

        print(f"[{idx + 1}/{len(windows)}] {label}")
        open_explorer_window(path, position=position, dry_run=dry_run)

        if idx < len(windows) - 1 and delay_ms > 0:
            time.sleep(delay_ms / 1000.0)

    print("[INFO] All windows launched.")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Windows File Explorer Manager – Startup Automation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python file_explorer_manager.py\n"
            "  python file_explorer_manager.py --config my_config.json\n"
            "  python file_explorer_manager.py --dry-run\n"
        ),
    )
    parser.add_argument(
        "--config",
        default=os.path.join(os.path.dirname(__file__), "config.json"),
        metavar="FILE",
        help="Path to the JSON configuration file (default: config.json)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be done without actually opening windows",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    run_from_config(args.config, dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    sys.exit(main())
