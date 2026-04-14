# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "pyautogui",
#     "pyperclip",
#     "pygetwindow",
# ]
# ///

import subprocess
import time
import pyautogui
import pyperclip
import pygetwindow as gw
import os
import json

# =================================================================
# MASTER SETTINGS
# =================================================================
# 1. SET TO False TO STOP THE SCRIPT FROM UPDATING THE STARTUP FOLDER
ENABLE_STARTUP_SYNC = True 

# 2. UNIQUE NAME: CHANGE THIS FOR EACH DIFFERENT SCRIPT/LIST 
# (e.g., "AutoExplorer_Work", "AutoExplorer_Personal")
STARTUP_IDENTITY = "AutoExplorer_Main" 

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "folders.json")
pyautogui.FAILSAFE = True

# Fallback folders if folders.json is missing
DEFAULT_FOLDERS = [
    os.path.expanduser("~/Downloads"),
    os.path.expanduser("~/Desktop"),
    r"C:\Windows\System32"
]

def load_config_robustly():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(DEFAULT_FOLDERS, f, indent=4)
        print(f"[CONFIG] Created new config with defaults at {CONFIG_FILE}")
        return DEFAULT_FOLDERS
    
    with open(CONFIG_FILE, 'r') as f:
        content = f.read()
        try:
            return json.load(f)
        except:
            # Fixes bad backslashes
            fixed = content.replace('\\', '\\\\').replace('\\\\\\\\', '\\\\')
            return json.loads(fixed)

def setup_windows_startup():
    """Links this specific script identity to the Windows Startup folder."""
    startup_folder = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")
    
    # We use STARTUP_IDENTITY here so multiple scripts can coexist
    vbs_path = os.path.join(startup_folder, f"{STARTUP_IDENTITY}.vbs")
    script_path = os.path.abspath(__file__)
    
    vbs_content = f'Set WshShell = CreateObject("WScript.Shell")\nWshShell.Run "uvw run """"{script_path}"""""", 0, False\n'
    
    with open(vbs_path, 'w') as f:
        f.write(vbs_content)
    print(f"[SETUP] Startup identity '{STARTUP_IDENTITY}' synced to: {script_path}")

def startup_explorer():
    raw_paths = load_config_robustly()
    valid_paths = [os.path.normpath(p) for p in raw_paths if os.path.exists(p)]
    
    if not valid_paths:
        print("[ERROR] No valid paths found.")
        return

    # 1. Minimize Terminal
    try:
        active_win = gw.getActiveWindow()
        if active_win and ("PowerShell" in active_win.title or "Command Prompt" in active_win.title):
            active_win.minimize()
    except:
        pass

    # 2. Launch Base Window
    print(f"[RUNNING] Launching {valid_paths[0]}...")
    subprocess.Popen(['explorer.exe', valid_paths[0]])
    
    # 3. Discover Window
    explorer_window = None
    for _ in range(15):
        time.sleep(0.5)
        wins = [w for w in gw.getAllWindows() if "File Explorer" in w.title or "Explor" in w.title]
        if wins:
            explorer_window = wins[0]
            break

    if not explorer_window:
        print("[ERROR] Could not find Explorer window.")
        return

    # 4. Force Focus
    try:
        explorer_window.restore()
        explorer_window.activate()
    except:
        width, height = pyautogui.size()
        pyautogui.click(width // 2, height // 2)
    
    time.sleep(3.0)

    # 5. Tab Injection
    for path in valid_paths[1:]:
        pyautogui.click(explorer_window.left + 150, explorer_window.top + 10)
        pyautogui.hotkey('ctrl', 't')
        time.sleep(1.5) 
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.9)
        
        pyperclip.copy(path)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.7)
        
        pyautogui.press('enter')
        time.sleep(0.9)
        
    print(f"[SUCCESS] {len(valid_paths)} tabs opened for identity: {STARTUP_IDENTITY}")

if __name__ == "__main__":
    # ONLY SYNC IF THE FLAG IS SET TO TRUE
    if ENABLE_STARTUP_SYNC:
        setup_windows_startup()
    else:
        print("[INFO] Startup sync is DISABLED. Running script manually.")

    try:
        startup_explorer()
    except Exception as e:
        print(f"\n[ERROR] Automation failed: {e}")