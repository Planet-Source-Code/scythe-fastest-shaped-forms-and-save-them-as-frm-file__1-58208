Attribute VB_Name = "Cmdialog"
'Subclassing Common Dialog
'This modul calls all needed common dialogs for Load / Save lmf´s and
'adds frmcmdlg to Load LMF so we can preview the logo
'I could include colordialog....
'but i havnt the time to do this (at the moment)
Option Explicit
'Open File Dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long
'Open/Save FileName Flags
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_EXPLORER = &H80000
Private Const OFN_OVERWRITEPROMPT = &H2
' Filetype
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Colordialog
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As ChooseColorType) As Long
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Type ChooseColorType
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_RGBINIT = &H1
'Windows Messages for Hook
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Private Const WM_INITDIALOG = &H110
'Common Dialog Messages
Private Const CDM_GETFILEPATH = &H465
Private Const CDM_GETFOLDERPATH = &H466
'Get Dialogs Position and Size
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Move/Resize the Dialog and our own form
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Find the selectet file
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
'This Api is used for fast array operations in LogoMan, Cmdialog, Compress
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private CdlgHwnd As Long 'This will hold the Hwnd for Common Dialog
'Open the common dialog (standard api way)
Public Function OpenDialog(FrmHwnd As Long, Startpath As String, Filter As String, Optional Prev As Boolean = False) As String

Dim Filebox As OPENFILENAME
Dim Result As Long

    With Filebox
        .lStructSize = Len(Filebox)
        .hwndOwner = FrmHwnd
        .hInstance = 0
        .lpstrFilter = Filter
        .nMaxCustomFilter = 0
        .nFilterIndex = 1
        .lpstrFile = Space(256) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = Startpath
'**Removed Prview because there will be no Preview in this Program
'**This feature was used in LogoMan
'        If Prev = True Then
'to use hook we have to set the OFN_ENABLEHOOK flag
'            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_ENABLEHOOK Or OFN_EXPLORER
'Give cmdlg the hook for our preview window
'            .lpfnHook = GetHookAdress(AddressOf CmdlgHook)
'            Else
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
'Give cmdlg the hook for our preview window
            .lpfnHook = 0
'        End If
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
    End With
    Result = GetOpenFileName(Filebox)
    If Result <> 0 Then
        OpenDialog = Left(Filebox.lpstrFile, InStr(Filebox.lpstrFile, vbNullChar) - 1)
    End If

End Function
'Save dialog
'Normal Api (nothing spezial)
Public Function SaveDialog(FrmHwnd As Long, Startpath As String, Filter As String) As String

Dim Filebox As OPENFILENAME
Dim Result As Long

    With Filebox
        .lStructSize = Len(Filebox)
        .hwndOwner = FrmHwnd
        .hInstance = 0
        .lpstrFilter = Filter
        .nMaxCustomFilter = 0
        .nFilterIndex = 1
        .lpstrFile = Space(256) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = Startpath
        .flags = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
        .nFileOffset = 0
        .nFileExtension = 0
'get the standard fileextension
'common dialog should automatic add to the filename
        .lpstrDefExt = Mid$(Filter, InStr(1, Filter, "*", vbBinaryCompare) + 2, 3)
        .lCustData = 0
        .lpfnHook = 0
    End With
    Result = GetSaveFileName(Filebox)
    If Result <> 0 Then
        SaveDialog = Left(Filebox.lpstrFile, InStr(Filebox.lpstrFile, vbNullChar) - 1)
    End If

End Function

Public Function ColorDialog(ByVal FrmHwnd As Long, Optional Colr As Long = 0) As Long

Dim Col As ChooseColorType
Dim Addr As Long
Dim Memh As Long
Dim ClrArray(15) As Long
Dim I As Long
Dim Result As Long
'Reserve Memory to hold our custom colors
    Memh = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 64)
    Addr = GlobalLock(Memh)
'Fill our Array white
    For I = 0 To UBound(ClrArray)
        ClrArray(I) = &HFFFFFF
    Next I
'This is only in cause i want them white
'Copy Array to Memory
    CopyMemory ByVal Addr, ClrArray(0), 64
    Col.lStructSize = Len(Col)
    Col.hwndOwner = FrmHwnd
    Col.lpCustColors = Addr
    Col.rgbResult = Colr
    Col.flags = CC_RGBINIT Or CC_FULLOPEN
    Result = ChooseColor(Col)
'Dont need this part cause we only want the selected color
'Get Colors Back
'CopyMemory ClrArray(0), ByVal Addr, 64
    GlobalUnlock Memh
    GlobalFree Memh
    If Result = 0 Then
        ColorDialog = -1
        Else
        ColorDialog = Col.rgbResult
    End If

End Function
