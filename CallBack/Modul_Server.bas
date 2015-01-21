Attribute VB_Name = "Modul_Server"
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
Public Port As String ' Начальный порт

' Общие настройки
Public Check_Url As String ' Проверяемый URL
Public Crypt_Key As String ' Ключ шифрования

Sub Main()
 Call InitXPStyle
 If App.PrevInstance Then
  MsgBox "Программа CallBack - Server уже запущена!", vbCritical, "Ошибка!!!"
  End
 End If
 Main_Server.Show
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

Public Sub Save_Setting()
 Dim i As Long
 Open App.Path & "\Setting.dat" For Output As #1
  Write #1, Port, Check_Url, Crypt_Key
 Close #1
End Sub

Public Function RC4(data As String) As String
 Dim Key(0 To 255) As Long
 Dim box(0 To 255) As Long
 Dim cipher As String
 Dim pwd_length As Long
 Dim data_length As Long
 Dim J As Long
 Dim a As Long
 Dim i As Long
 Dim k As Long
 Dim Tmp As Long
 Dim txt As String
 Dim pwd As String
 pwd = Crypt_Key
 pwd_length = Len(pwd)
 data_length = Len(data)
 For i = 0 To 255
  Key(i) = Asc(Mid$(pwd, (i Mod pwd_length) + 1, 1))
  box(i) = i
 Next
 J = 0
 For i = 0 To 255
  J = (J + box(i) + Key(i)) Mod 256
  Tmp = box(i)
  box(i) = box(J)
  box(J) = Tmp
 Next
 a = 0
 J = 0
 For i = 1 To data_length
  a = (a + 1) Mod 256
  J = (J + box(a)) Mod 256
  Tmp = box(a)
  box(a) = box(J)
  box(J) = Tmp
  k = box((box(a) + box(J)) Mod 256)
  cipher = cipher & Chr$(Asc(Mid$(data, i, 1)) Xor k)
 Next
 RC4 = cipher
End Function

Function RC4_DeCode(txt As String) As String
 Dim s As String
 Dim c As String
 Dim rs1 As String
 Dim rs2 As String
 Dim r1 As Long
 Dim r2 As Long
 Dim i As Long
 For i = 1 To Len(txt) Step 2
  rs1 = Mid$(txt, i, 1)
  rs2 = Mid$(txt, i + 1, 1)
  If Asc(rs1) - 48 > -1 And Asc(rs1) - 48 < 10 Then r1 = Asc(rs1) - 48
  If Asc(rs1) - 55 > 9 And Asc(rs1) - 55 < 16 Then r1 = Asc(rs1) - 55
  If Asc(rs2) - 48 > -1 And Asc(rs2) - 48 < 10 Then r2 = Asc(rs2) - 48
  If Asc(rs2) - 55 > 9 And Asc(rs2) - 55 < 16 Then r2 = Asc(rs2) - 55
  s = s & Chr$(r1 * 16 + r2)
 Next
 RC4_DeCode = s
End Function

Function RC4_EnCode(txt As String) As String
 Dim s As String
 Dim c As String
 Dim i As Long
 For i = 1 To Len(txt)
  c = Hex$(Asc(Mid$(txt, i, 1)))
  If Len(c) = 1 Then c = "0" & c
  s = s & c
 Next
 RC4_EnCode = s
End Function

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
 ShellExecute Main_Server.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub

Public Sub Add_Status(i As Integer, txt As String)
 Dim Log_Size As Long
 Dim Str As String
 Str = IIf(i <> 0, i & " - ", "") & txt
 Main_Server.Label_Status.Caption = Str
 If Not IsEmpty(Logi) Then
  Logi.List_Log.AddItem Format(Date, "dd.mm.yyyy") & " " & Format(Time, "(HH:mm:ss)") & ": " & Str
  Logi.List_Log.ListIndex = Logi.List_Log.ListCount - 1
  Call Main_Server.ClientCount
 End If
 DoEvents
 Open App.Path & "\Logs\" & Format(Date, "yyyy-mm-dd") & ".log" For Append As #1
  Print #1, Format(Date, "dd.mm.yyyy") & " " & Format(Time, "(HH:mm:ss)") & ": " & Str
  Log_Size = LOF(1)
 Close #1
End Sub
