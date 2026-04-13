"""
install_startup.py
------------------
Installs (or removes) a shortcut to startup_launcher.bat in the Windows
Startup folder so that the File Explorer automation runs on every login.

Usage:
  python install_startup.py install    -- add to startup
  python install_startup.py remove     -- remove from startup
  python install_startup.py status     -- check if currently installed

Requirements: Python 3.8+, Windows OS
"""

import argparse
import os
import sys
import shutil


STARTUP_FOLDER = os.path.join(
    os.environ.get("APPDATA", ""),
    "Microsoft",
    "Windows",
    "Start Menu",
    "Programs",
    "Startup",
)

SHORTCUT_NAME = "FileExplorerManager.bat"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LAUNCHER_SRC = os.path.join(SCRIPT_DIR, "startup_launcher.bat")
SHORTCUT_DEST = os.path.join(STARTUP_FOLDER, SHORTCUT_NAME)


def _check_windows() -> None:
    if sys.platform != "win32":
        print("[ERROR] This script only works on Windows.")
        sys.exit(1)


def cmd_install() -> None:
    """Copy startup_launcher.bat to the Windows Startup folder."""
    _check_windows()

    if not os.path.isfile(LAUNCHER_SRC):
        print(f"[ERROR] Launcher not found: {LAUNCHER_SRC}")
        sys.exit(1)

    os.makedirs(STARTUP_FOLDER, exist_ok=True)
    shutil.copy2(LAUNCHER_SRC, SHORTCUT_DEST)
    print(f"[OK] Installed startup shortcut:\n     {SHORTCUT_DEST}")
    print("     The File Explorer Manager will now run on every login.")


def cmd_remove() -> None:
    """Remove the startup shortcut if it exists."""
    _check_windows()

    if os.path.isfile(SHORTCUT_DEST):
        os.remove(SHORTCUT_DEST)
        print(f"[OK] Removed startup shortcut:\n     {SHORTCUT_DEST}")
    else:
        print("[INFO] Startup shortcut is not installed – nothing to remove.")


def cmd_status() -> None:
    """Report whether the startup shortcut is currently installed."""
    _check_windows()

    if os.path.isfile(SHORTCUT_DEST):
        print(f"[INSTALLED] Startup shortcut found:\n            {SHORTCUT_DEST}")
    else:
        print("[NOT INSTALLED] Startup shortcut is not present.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Install or remove the File Explorer Manager from Windows Startup",
    )
    parser.add_argument(
        "command",
        choices=["install", "remove", "status"],
        help="Action to perform",
    )
    args = parser.parse_args(argv)

    dispatch = {
        "install": cmd_install,
        "remove": cmd_remove,
        "status": cmd_status,
    }
    dispatch[args.command]()
    return 0


if __name__ == "__main__":
    sys.exit(main())
