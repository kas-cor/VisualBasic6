Attribute VB_Name = "Tray"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" _
            Alias "Shell_NotifyIconA" _
            (ByVal dwMessage As Long, _
            pnid As NOTIFYICONDATA) As Boolean
            
Dim TrayI As NOTIFYICONDATA

Public Sub Tray_Add(Picture As Long, Tip As String, hwnd As Long)
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = hwnd
    TrayI.uId = vbNull
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = Picture
    TrayI.szTip = Tip & vbNullString
    Shell_NotifyIcon NIM_ADD, TrayI
End Sub

Public Sub Tray_Del(hwnd As Long)
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = hwnd
    TrayI.uId = vbNull
    Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

Public Sub Tray_Modify(Picture As Long)
    TrayI.hIcon = Picture
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Public Sub Tray_Modify_Tip(Tip As String)
    TrayI.szTip = Tip & vbNullString
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Sub File_Run(File As String)
 ShellExecute Main.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub
