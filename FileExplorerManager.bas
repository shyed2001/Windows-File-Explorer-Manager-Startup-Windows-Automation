Attribute VB_Name = "FileExplorerManager"
'
' FileExplorerManager.bas
' -----------------------
' VBA module: Windows File Explorer Manager – Startup Windows Automation
'
' Provides macros to open, position, and manage multiple File Explorer
' windows from Excel or any Office application that hosts VBA.
'
' Import this module into an Excel workbook:
'   Developer tab > Visual Basic > File > Import File > FileExplorerManager.bas
'
' Then call the macros from a button or the Macros dialog (Alt+F8).
'
' Tested on: Windows 10/11, Office 2016+
'

Option Explicit

' ---------------------------------------------------------------------------
' Win32 API declarations
' ---------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
         ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal uFlags As Long) As Long

    Private Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

    Private Declare PtrSafe Function GetTopWindow Lib "user32" _
        (ByVal hWnd As LongPtr) As LongPtr

    Private Declare PtrSafe Function GetWindow Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr

    Private Declare PtrSafe Function GetWindowTextW Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal uFlags As Long) As Long

    Private Declare Function ShowWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

    Private Declare Function GetTopWindow Lib "user32" _
        (ByVal hWnd As Long) As Long

    Private Declare Function GetWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal uCmd As Long) As Long

    Private Declare Function GetWindowTextW Lib "user32" _
        (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#End If

' ---------------------------------------------------------------------------
' Constants
' ---------------------------------------------------------------------------

Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SW_RESTORE        As Long = 9
Private Const GW_HWNDNEXT       As Long = 2

' ---------------------------------------------------------------------------
' Public macros – call these from the Macros dialog (Alt+F8)
' ---------------------------------------------------------------------------

'
' OpenDefaultWindows
' ------------------
' Opens the three most commonly used File Explorer windows: Desktop,
' Documents, and Downloads, each positioned side-by-side.
'
Public Sub OpenDefaultWindows()
    OpenExplorerWindow Environ("USERPROFILE") & "\Desktop",    0,   0, 800, 600
    Wait 300
    OpenExplorerWindow Environ("USERPROFILE") & "\Documents",  820, 0, 800, 600
    Wait 300
    OpenExplorerWindow Environ("USERPROFILE") & "\Downloads",  0, 620, 800, 600

    MsgBox "File Explorer windows opened successfully.", vbInformation, "File Explorer Manager"
End Sub

'
' OpenCustomWindow
' ----------------
' Prompts the user for a folder path and opens it in File Explorer.
'
Public Sub OpenCustomWindow()
    Dim folderPath As String
    folderPath = InputBox("Enter the full folder path to open:", "Open File Explorer", _
                          Environ("USERPROFILE"))

    If Len(Trim(folderPath)) = 0 Then
        MsgBox "No path entered. Operation cancelled.", vbExclamation, "File Explorer Manager"
        Exit Sub
    End If

    OpenExplorerWindow folderPath, 100, 100, 900, 650
End Sub

'
' OpenFromSheet
' -------------
' Reads folder paths and window positions from the active worksheet and opens
' them all.  Expected columns:
'   A = Folder path
'   B = Left   (pixels, default 0)
'   C = Top    (pixels, default 0)
'   D = Width  (pixels, default 800)
'   E = Height (pixels, default 600)
'
' Row 1 is treated as a header row and skipped.
'
Public Sub OpenFromSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim folderPath As String
    Dim posLeft As Long, posTop As Long, posWidth As Long, posHeight As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found.  Please add folder paths starting from row 2.", _
               vbExclamation, "File Explorer Manager"
        Exit Sub
    End If

    Dim opened As Long
    opened = 0

    For i = 2 To lastRow
        folderPath = Trim(ws.Cells(i, "A").Value)
        If Len(folderPath) = 0 Then GoTo NextRow

        posLeft   = IIf(IsNumeric(ws.Cells(i, "B").Value), CLng(ws.Cells(i, "B").Value), 0)
        posTop    = IIf(IsNumeric(ws.Cells(i, "C").Value), CLng(ws.Cells(i, "C").Value), 0)
        posWidth  = IIf(IsNumeric(ws.Cells(i, "D").Value) And ws.Cells(i, "D").Value > 0, _
                        CLng(ws.Cells(i, "D").Value), 800)
        posHeight = IIf(IsNumeric(ws.Cells(i, "E").Value) And ws.Cells(i, "E").Value > 0, _
                        CLng(ws.Cells(i, "E").Value), 600)

        OpenExplorerWindow folderPath, posLeft, posTop, posWidth, posHeight
        opened = opened + 1
        Wait 300

NextRow:
    Next i

    MsgBox opened & " File Explorer window(s) opened.", vbInformation, "File Explorer Manager"
End Sub

'
' CloseAllExplorerWindows
' -----------------------
' Closes every open File Explorer window.
'
Public Sub CloseAllExplorerWindows()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Close ALL open File Explorer windows?", _
                    vbQuestion + vbYesNo, "File Explorer Manager")
    If answer = vbNo Then Exit Sub

    Shell "taskkill /F /IM explorer.exe /T", vbHide
    Wait 1500
    ' Restart the Windows shell so that the taskbar is restored.
    Shell "explorer.exe", vbNormalFocus
    MsgBox "All File Explorer windows have been closed and the shell restarted.", _
           vbInformation, "File Explorer Manager"
End Sub

' ---------------------------------------------------------------------------
' Private helpers
' ---------------------------------------------------------------------------

'
' OpenExplorerWindow
' ------------------
' Opens a File Explorer window at the given path and positions/sizes it.
'
Private Sub OpenExplorerWindow(ByVal folderPath As String, _
                                ByVal posLeft As Long, ByVal posTop As Long, _
                                ByVal posWidth As Long, ByVal posHeight As Long)

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MsgBox "Path not found, skipping:" & vbCrLf & folderPath, _
               vbExclamation, "File Explorer Manager"
        Exit Sub
    End If

    Shell "explorer.exe """ & folderPath & """", vbNormalFocus

    ' Give Explorer time to create its window before we try to find it.
    Wait 800

    Dim hWnd As LongPtr
    hWnd = FindExplorerWindow(folderPath)

#If VBA7 Then
    If hWnd <> 0 Then
        ShowWindow hWnd, SW_RESTORE
        SetWindowPos hWnd, 0, posLeft, posTop, posWidth, posHeight, _
                     SWP_NOZORDER Or SWP_NOACTIVATE
    End If
#Else
    If hWnd <> 0 Then
        ShowWindow hWnd, SW_RESTORE
        SetWindowPos hWnd, 0, posLeft, posTop, posWidth, posHeight, _
                     SWP_NOZORDER Or SWP_NOACTIVATE
    End If
#End If
End Sub

'
' FindExplorerWindow
' ------------------
' Walks all top-level windows and returns the HWND whose title contains the
' last folder component of folderPath.  Returns 0 if not found.
'
#If VBA7 Then
Private Function FindExplorerWindow(ByVal folderPath As String) As LongPtr
    Dim folderName As String
    folderName = LCase(GetFolderName(folderPath))

    Dim hWnd As LongPtr
    Dim titleBuf As String
    Dim titleLen As Long

    hWnd = GetTopWindow(0)
    Do While hWnd <> 0
        titleBuf = String(512, Chr(0))
        titleLen = GetWindowTextW(hWnd, titleBuf, 512)
        If titleLen > 0 Then
            If InStr(1, LCase(Left(titleBuf, titleLen)), folderName) > 0 Then
                FindExplorerWindow = hWnd
                Exit Function
            End If
        End If
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop

    FindExplorerWindow = 0
End Function
#Else
Private Function FindExplorerWindow(ByVal folderPath As String) As Long
    Dim folderName As String
    folderName = LCase(GetFolderName(folderPath))

    Dim hWnd As Long
    Dim titleBuf As String
    Dim titleLen As Long

    hWnd = GetTopWindow(0)
    Do While hWnd <> 0
        titleBuf = String(512, Chr(0))
        titleLen = GetWindowTextW(hWnd, titleBuf, 512)
        If titleLen > 0 Then
            If InStr(1, LCase(Left(titleBuf, titleLen)), folderName) > 0 Then
                FindExplorerWindow = hWnd
                Exit Function
            End If
        End If
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop

    FindExplorerWindow = 0
End Function
#End If

'
' GetFolderName
' -------------
' Returns the last component of a path string (the folder name).
'
Private Function GetFolderName(ByVal folderPath As String) As String
    Dim parts() As String
    Dim cleaned As String
    cleaned = folderPath
    If Right(cleaned, 1) = "\" Or Right(cleaned, 1) = "/" Then
        cleaned = Left(cleaned, Len(cleaned) - 1)
    End If
    parts = Split(cleaned, "\")
    GetFolderName = parts(UBound(parts))
End Function

'
' Wait
' ----
' Pauses execution for the specified number of milliseconds.
'
Private Sub Wait(ByVal milliseconds As Long)
    Dim endTime As Date
    endTime = Now + milliseconds / 86400000#
    Do While Now < endTime
        DoEvents
    Loop
End Sub
