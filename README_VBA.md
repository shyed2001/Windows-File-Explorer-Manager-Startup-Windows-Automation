# VBA Module – How to Import and Use FileExplorerManager.bas

This guide explains how to import the `FileExplorerManager.bas` VBA module into
Microsoft Excel (or any Office application) and run its macros.

---

## Prerequisites

| Requirement | Details |
|---|---|
| Operating System | Windows 10 or Windows 11 |
| Office Version | Excel 2016, 2019, 2021, or Microsoft 365 |
| Macro Security | Must allow macros (see step 1 below) |

---

## Step 1 – Enable Macros in Excel

1. Open Excel.
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Select **Macro Settings** and choose **Enable all macros** (or *Enable macros with notification*).
4. Click **OK** twice.

---

## Step 2 – Open the VBA Editor

Press **Alt + F11** to open the Visual Basic for Applications (VBA) editor.

---

## Step 3 – Import the Module

1. In the VBA editor menu bar, click **File → Import File…**
2. Browse to the location where you saved `FileExplorerManager.bas`.
3. Select the file and click **Open**.
4. You should now see a **FileExplorerManager** module in the *Modules* folder
   in the Project Explorer on the left.

---

## Step 4 – Run a Macro

### Using the Macros dialog (Alt + F8)

1. Press **Alt + F8** in Excel.
2. Select one of the macros from the list:

| Macro | Description |
|---|---|
| `OpenDefaultWindows` | Opens Desktop, Documents, and Downloads side-by-side |
| `OpenCustomWindow` | Prompts for a folder path and opens it |
| `OpenFromSheet` | Reads paths and positions from the active worksheet |
| `CloseAllExplorerWindows` | Closes every open File Explorer window |

3. Click **Run**.

### Using a Button

1. Go to the **Developer** tab (if not visible: File → Options → Customize Ribbon → tick Developer).
2. Click **Insert → Button (Form Control)**.
3. Draw the button on your sheet.
4. In the *Assign Macro* dialog, select `OpenDefaultWindows` (or another macro).
5. Click **OK**, then click the button to launch Explorer windows.

---

## Step 5 – Use OpenFromSheet (Optional)

The `OpenFromSheet` macro reads configuration from the **active worksheet**.
Set up your sheet like this:

| A (Path) | B (Left) | C (Top) | D (Width) | E (Height) |
|---|---|---|---|---|
| C:\Users\You\Desktop | 0 | 0 | 800 | 600 |
| C:\Users\You\Documents | 820 | 0 | 800 | 600 |
| C:\Projects | 0 | 620 | 1000 | 600 |

- Row 1 is the **header row** and is skipped.
- Columns B–E are optional; defaults are Left=0, Top=0, Width=800, Height=600.

---

## Saving the Workbook

To keep the macros, save as **Excel Macro-Enabled Workbook (*.xlsm)**.
If you save as `.xlsx`, all VBA code will be removed by Excel.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Cannot run the macro" | Re-import the `.bas` file; check macro security settings |
| Window not repositioned | Increase the `Wait` delay in the module (default 800 ms) |
| "Path not found" error | Verify the folder path exists and is accessible |
| Explorer restarts after close | This is normal – the Windows shell is restarted automatically |
