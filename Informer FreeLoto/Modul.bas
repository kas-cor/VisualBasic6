Attribute VB_Name = "Modul"
Option Explicit

Const Table As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Public Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (ByVal lpSemaphoreAttributes As Long, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
Public Declare Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long

Global semHNDL As Long

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As SINITCOMMONCONTROLSEX) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Type SINITCOMMONCONTROLSEX
   dwSize As Long
   dwICC As Long
End Type

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_NULL = &H0

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Var_Bilet As Integer

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

Sub Main()
 Call InitXPStyle
 Err.Clear
 semHNDL = 0
 semHNDL = CreateSemaphore(0, 0, 1, "InfoFreeLoto")
 If (Err.LastDllError <> 0) Or (semHNDL = 0) Then
  '::: Это - не первый экземпляр
  MsgBox "Программа Informer FreeLoto уже запущена!", vbCritical, "Ошибка!"
  End
 End If
 frmMain.Show
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
 ShellExecute frmMain.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub

Function AsciiToBase64(sSource As String) As String
    Dim t As Integer
    Dim l As Integer
    Dim m As Integer
    Dim Bin As String, Pos As Integer
    For t = 1 To Len(sSource) Step 3
    Bin = ""
    For l = t To t + 2
    If l > Len(sSource) Then
    Bin = Bin & "0000000="
    Else
    Bin = Bin & GetBin(Asc(Mid$(sSource, l, 1)), 8)
    End If
    Next l
    For m = 1 To Len(Bin) Step 6
    Pos = GetDec(Mid$(Bin, m, 6))
    AsciiToBase64 = AsciiToBase64 & Mid$(Table, Pos + 1, 1)
    Next m
    Next t
End Function

Function Base64ToAscii(sSource As String) As String
    Dim Bin As String, Pos As Integer, sLen As Long, sKrt As Long, dBrake As Long
    Dim t As Integer
    Dim l As Integer
    Dim m As Integer
    sLen = Len(sSource)
    For t = 1 To Len(sSource) Step 4
    Bin = ""
    For l = t To t + 3
    Pos = InStr(1, Table, Mid$(sSource, l, 1))
    Bin = Bin & GetBin(CStr(Pos - 1), 6)
    Next l
    For m = 1 To Len(Bin) Step 8
    If GetDec(Mid$(Bin, m, 8)) <> 256 Then
    Base64ToAscii = Base64ToAscii & Chr$(GetDec(Mid$(Bin, m, 8)))
    End If
    Next m
    Next t
End Function

Function GetBin(Dec As Single, cFormat As Integer) As String
    Dim sFormat As String
    sFormat = String(cFormat, "0")
    If Dec = 64 Then GetBin = "00000 ": Exit Function
    If Dec = 0 Then GetBin = sFormat: Exit Function
    Do While Dec >= 1
    Dec = Int(Dec) / 2
    If Dec = Int(Dec) Then
    GetBin = 0 & GetBin
    Else
    GetBin = 1 & GetBin
    End If
    Loop
    GetBin = Format$(GetBin, sFormat)
End Function

Function GetDec(Bin As String) As Long
    Dim t As Integer
    If InStr(1, Bin, "=") Then GetDec = 64: Exit Function
    If InStr(1, Bin, " ") Then GetDec = 256: Exit Function
    Dim Cnt As Integer
    Cnt = Len(Bin) - 1
    For t = 1 To Len(Bin)
    If Mid(Bin, t, 1) = 1 Then
    GetDec = GetDec + (2 ^ Cnt)
    End If
    Cnt = Cnt - 1
    Next
End Function

Public Function RC4(data As String) As String
 Dim Key(0 To 255) As Long
 Dim box(0 To 255) As Long
 Dim cipher As String
 Dim pwd_length As Long
 Dim data_length As Long
 Dim J As Long
 Dim a As Long
 Dim i As Long
 Dim K As Long
 Dim Tmp As Long
 Dim Txt As String
 Dim pwd As String
 pwd = "{password}"
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
  K = box((box(a) + box(J)) Mod 256)
  cipher = cipher & Chr(Asc(Mid$(data, i, 1)) Xor K)
 Next
 RC4 = cipher
End Function

Function FrLt_DeCode(Txt As String) As String
 Dim s As String
 Dim c As String
 Dim rs1 As String
 Dim rs2 As String
 Dim r1 As Long
 Dim r2 As Long
 Dim i As Long
 For i = 1 To Len(Txt) Step 2
  rs1 = Mid$(Txt, i, 1)
  rs2 = Mid$(Txt, i + 1, 1)
  If Asc(rs1) - 48 > -1 And Asc(rs1) - 48 < 10 Then r1 = Asc(rs1) - 48
  If Asc(rs1) - 55 > 9 And Asc(rs1) - 55 < 16 Then r1 = Asc(rs1) - 55
  If Asc(rs2) - 48 > -1 And Asc(rs2) - 48 < 10 Then r2 = Asc(rs2) - 48
  If Asc(rs2) - 55 > 9 And Asc(rs2) - 55 < 16 Then r2 = Asc(rs2) - 55
  s = s & Chr$(r1 * 16 + r2)
 Next
 FrLt_DeCode = s
End Function

Function FrLt_EnCode(Txt As String) As String
 Dim s As String
 Dim c As String
 Dim i As Long
 For i = 1 To Len(Txt)
  c = Hex$(Asc(Mid$(Txt, i, 1)))
  If Len(c) = 1 Then c = "0" & c
  s = s & c
 Next
 FrLt_EnCode = s
End Function

Public Function Pass_DeCrypt(Pass As String) As String
 If Mid$(Pass, 1, 6) = "Crypt:" Then
  Pass_DeCrypt = FrLt_DeCode(Mid$(Pass, 7))
 Else
  Pass_DeCrypt = Pass
 End If
End Function

Public Function Pass_EnCrypt(Pass As String) As String
 Pass_EnCrypt = "Crypt:" & FrLt_EnCode(Pass)
End Function
