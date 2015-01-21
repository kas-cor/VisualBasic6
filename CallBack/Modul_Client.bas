Attribute VB_Name = "Modul_Client"
Option Explicit

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As SINITCOMMONCONTROLSEX) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Type SINITCOMMONCONTROLSEX
   dwSize As Long
   dwICC As Long
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_NULL = &H0

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

Public Declare Function Shell_NotifyIcon Lib "shell32" _
            Alias "Shell_NotifyIconA" _
            (ByVal dwMessage As Long, _
            pnid As NOTIFYICONDATA) As Boolean

Public TrayI As NOTIFYICONDATA

' Настройки сети
Public Port As String '             Начальный порт
Public Host As String '             Адрес сервера
Public Manager As String '          имя менеджера

Public Data As String '             Данные

Public Message As String '          Сообщение о клиенте

Public Op_WinSignal As Boolean '    Открыто окно сигнала
Public Op_WinInfo As Boolean '      Открыто окно информации
Public Cl_WinSignal As Boolean '    Закрытите окна сигнала
Public Cl_WinSignal_Men As String ' Закрытите окна сигнала (менеджер)

Sub Main()
 Call InitXPStyle
 If App.PrevInstance Then
  MsgBox "Программа CallBack уже запущена!", vbCritical, "Ошибка!!!"
  End
 End If
 Main_Client.Show
End Sub

Public Sub InitXPStyle()
 Dim InitCtrls As SINITCOMMONCONTROLSEX
 On Error Resume Next
 With InitCtrls
  .dwSize = LenB(InitCtrls)
  .dwICC = ICC_USEREX_CLASSES
 End With
 InitCommonControlsEx InitCtrls
End Sub

Public Sub Tray_Add(Picture As Long, Tip As String, hwnd As Long)
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = hwnd
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = Picture
    TrayI.szTip = Tip & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayI
End Sub

Public Sub Tray_Del(hwnd As Long)
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

Public Sub Tray_Modify(Picture As Long)
    TrayI.hIcon = Picture
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Public Sub Tray_Modify_Tip(Tip As String)
    TrayI.szTip = Tip & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Sub File_Run(File As String)
 ShellExecute Main_Client.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub

Public Sub Add_Status(txt As String)
 Main_Client.Label_Status.Caption = txt
 Main_Client.Label_Status.ToolTipText = txt
 DoEvents
End Sub

Public Function SendCom(Com As String) As String
 On Error Resume Next
 If Main_Client.Winsock.State = sckClosed Then
  SendCom = "Error"
  Exit Function
 End If
 Dim Tmr As Long
 Dim Flag_Err As Boolean
 Data = vbNullString
 Main_Client.Winsock.SendData Com & "." & vbCrLf
 Tmr = Timer()
 Do
  DoEvents
  If Timer() - Tmr > 60 Then
   Add_Status "Таймаут при попытке получить данные"
   Flag_Err = True
  End If
 Loop Until InStr(1, Data, "." & vbCrLf) <> 0 Or Flag_Err
 If Not Flag_Err Then
  SendCom = Mid$(Data, 1, Len(Data) - 3)
 Else
  SendCom = "Error"
 End If
End Function

