# Windows File Explorer Manager – Startup Automation

Automatically open and position multiple **Windows File Explorer** windows
when your PC starts, using either a **Python** script or a **VBA** macro in
Microsoft Excel.

---

## Features

- Open any number of File Explorer windows at startup
- Position and resize each window precisely on screen
- Configuration-driven – edit a single JSON file to change folders/layouts
- One-click installer that adds the launcher to the Windows Startup folder
- VBA module for the same functionality directly from Excel

---

## Project Layout

```
├── config.json                  # Folder paths and window positions
├── file_explorer_manager.py     # Python automation script
├── startup_launcher.bat         # Batch file that runs the Python script
├── install_startup.py           # Adds/removes the launcher from Windows Startup
├── FileExplorerManager.bas      # VBA module (import into Excel)
├── README_VBA.md                # Step-by-step VBA import guide
└── test_file_explorer_manager.py  # Unit tests (pytest)
```

---

## Quick Start – Python

### Requirements

- Windows 10 or Windows 11
- Python 3.8 or later (standard library only – no extra packages needed)

### 1. Edit `config.json`

Open `config.json` and set the folders you want to open at startup:

```json
{
  "windows": [
    {
      "path": "C:\\Users\\%USERNAME%\\Desktop",
      "label": "Desktop",
      "position": { "left": 0, "top": 0, "width": 800, "height": 600 }
    },
    {
      "path": "C:\\Users\\%USERNAME%\\Documents",
      "label": "Documents",
      "position": { "left": 820, "top": 0, "width": 800, "height": 600 }
    }
  ],
  "settings": {
    "delay_between_windows_ms": 300
  }
}
```

`%USERNAME%` and other environment variables are expanded automatically.
The `position` key is optional; omit it to let Windows choose the placement.

### 2. Run manually

```bat
python file_explorer_manager.py
```

Preview what will happen without opening anything:

```bat
python file_explorer_manager.py --dry-run
```

Use a custom config file:

```bat
python file_explorer_manager.py --config my_config.json
```

### 3. Install at startup

```bat
python install_startup.py install
```

This copies `startup_launcher.bat` to your Windows Startup folder so it runs
automatically every time you log in.

To check whether it is installed:

```bat
python install_startup.py status
```

To remove it:

```bat
python install_startup.py remove
```

---

## Quick Start – VBA (Microsoft Excel)

See **[README_VBA.md](README_VBA.md)** for the full step-by-step guide.

**Short version:**

1. Open Excel and press **Alt + F11** to open the VBA editor.
2. **File → Import File…** → select `FileExplorerManager.bas`.
3. Press **Alt + F8** → run `OpenDefaultWindows`.

Available macros:

| Macro | Description |
|---|---|
| `OpenDefaultWindows` | Opens Desktop, Documents, Downloads side-by-side |
| `OpenCustomWindow` | Prompts for a folder path and opens it |
| `OpenFromSheet` | Reads paths and positions from the active worksheet |
| `CloseAllExplorerWindows` | Closes all open File Explorer windows |

---

## Running Tests

```bat
python -m pytest test_file_explorer_manager.py -v
```

The tests mock all Windows-specific calls so they run on any OS (including
Linux and macOS CI environments).

---

## License

[Apache 2.0](LICENSE)
