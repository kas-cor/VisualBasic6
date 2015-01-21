Attribute VB_Name = "Modul"
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

Public Crypt_Key As String * 1024 '       Ключ для шифрования

Public Users_Max As Long '                      Всего пользователей
Public Users_Login(0 To 1000) As String '       Логин (имя)
Public Users_Pass(0 To 1000) As String '        Пароль
Public Users_ModeStav(0 To 1000) As Boolean '   Режим рассчета ставки в час
Public Users_ZP(0 To 1000) As Double '          Оклад
Public Users_Stav(0 To 1000) As Double '        Ставка в час
Public Users_Edit_Time(0 To 1000) As Boolean '  Разрешать вводить время вручную
Public Users_Enter_Obed(0 To 1000) As Boolean ' Разрешать вводить время обеда
Public Users_Path(0 To 1000) As String '        Путь до базы данных пользователя

Public Data_Max As Long '                Всего записей
Public Data_Date(0 To 400) As String '   Дата
Public Data_Start(0 To 400) As Long '    Время прихода
Public Data_ObStart(0 To 400) As Long '  Время начало обеда
Public Data_ObEnd(0 To 400) As Long '    Время конца обеда
Public Data_End(0 To 400) As Long '      Время ухода

Public DontWork_Max As Long '             Всего не рабочих дней
Public DontWork_Date(0 To 1000) As Long ' Дата

' Общие настройки
Public Admin_Login As String '   Логин админа
Public Admin_Pass As String '    Пароль админа
Public Buh_Login As String '     Логин бухгалтера
Public Buh_Pass As String '      Пароль бухгалтера
Public Ports As String '         Порты
Public StartObed_Hr As Long '    Начало обеда (час.)
Public StartObed_Min As Long '   Начало обеда (мин.)
Public EndObed_Hr As Long '      Конец обеда (час.)
Public EndObed_Min As Long '     Конец обеда (мин.)
Public Work_Hr As Long '         Рабочих часов

' Сессии
Public SID_Status(0 To 1000) As Boolean ' Статус сессии
Public SID_Login(0 To 1000) As String '   Логин
Public SID_LogNum(0 To 1000) As String '  Номер пользователя
Public SID_Type(0 To 1000) As Long '      Тип пользователя 0-Пользователь, 1-бухгалтер, 2-администратор

Sub Main()
 Call InitXPStyle
 S_Main.Show
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
  Write #1, Admin_Login, Pass_EnCrypt(Admin_Pass), Buh_Login, Pass_EnCrypt(Buh_Pass), Ports, StartObed_Hr, StartObed_Min, EndObed_Hr, Work_Hr
 Close #1
End Sub

Public Sub Save_Users()
 Dim i As Long
 Open App.Path & "\Users.dat" For Output As #1
  Write #1, Users_Max
  For i = 0 To Users_Max - 1
   Write #1, Users_Login(i), Pass_EnCrypt(Users_Pass(i)), Users_ModeStav(i), Users_ZP(i), Users_Stav(i), Users_Edit_Time(i), Users_Enter_Obed(i), Users_Path(i)
  Next
 Close #1
End Sub

Public Sub Save_DontWork()
 Dim i As Long
 Open App.Path & "\DontWork.dat" For Output As #1
  Write #1, DontWork_Max
  For i = 0 To DontWork_Max - 1
   Write #1, DontWork_Date(i)
  Next
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

Function ZP_DeCode(txt As String) As String
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
 ZP_DeCode = s
End Function

Function ZP_EnCode(txt As String) As String
 Dim s As String
 Dim c As String
 Dim i As Long
 For i = 1 To Len(txt)
  c = Hex$(Asc(Mid$(txt, i, 1)))
  If Len(c) = 1 Then c = "0" & c
  s = s & c
 Next
 ZP_EnCode = s
End Function

Public Function Pass_DeCrypt(Pass As String) As String
 If Mid$(Pass, 1, 6) = "Crypt:" Then
  Pass_DeCrypt = ZP_DeCode(Mid$(Pass, 7))
 Else
  Pass_DeCrypt = Pass
 End If
End Function

Public Function Pass_EnCrypt(Pass As String) As String
 Pass_EnCrypt = "Crypt:" & ZP_EnCode(Pass)
End Function

Public Function Unix_Time(Hr As Long, Min As Long, Day As Long, Mount As Long, Year As Long) As Long
 Dim u As Long
 Dim Y As Long
 Y = Year
 If Y < 1970 Then Y = 0
 Y = Y - 1970
 u = Y * 599040 ' Минут в году
 u = u + Mount * 46080 ' Минут в месяце
 u = u + Day * 1440 ' Минут в сутках
 u = u + Hr * 60 ' Минут в часе
 u = u + Min
 Unix_Time = u
End Function

Public Function Time_Unix(u As Long) As String
 Dim Tmp As Long
 Dim Hr As Long
 Dim Hr_Zero As String
 Dim Min As Long
 Dim Min_Zero As String
 Dim Day As Long
 Dim Day_Zero As String
 Dim Mount As Long
 Dim Mount_Zero As String
 Dim Year As Long
 Tmp = u
 Year = Int(Tmp / 599040) ' Минут в году
 Tmp = Tmp - Year * 599040
 Mount = Int(Tmp / 46080) ' Минут в месяце
 Tmp = Tmp - Mount * 46080
 Day = Int(Tmp / 1440) ' Минут в дне
 Tmp = Tmp - Day * 1440
 Hr = Int(Tmp / 60) ' Минут в часе
 Tmp = Tmp - Hr * 60
 Min = Tmp
 If Hr < 10 Then Hr_Zero = "0"
 If Min < 10 Then Min_Zero = "0"
 If Day < 10 Then Day_Zero = "0"
 If Mount < 10 Then Mount_Zero = "0"
 Time_Unix = Hr_Zero & Hr & ":" & Min_Zero & Min & " " & Day_Zero & Day & "." & Mount_Zero & Mount & "." & Year + 1970
End Function

Public Sub Add_Status(i As Integer, txt As String)
 Dim Log_Size As Long
 Dim Str As String
 Str = IIf(i <> 0, i & " - ", "") & txt
 S_Main.StatusBar.Panels(1).Text = Str
 If Not IsEmpty(Form_Status) Then
  Form_Status.List_Log.AddItem Str
  Form_Status.List_Log.ListIndex = Form_Status.List_Log.ListCount - 1
  Call S_Main.ClientCount
 End If
 DoEvents
 Open App.Path & "\log.log" For Append As #1
  Print #1, Format(Date, "dd.mm.yyyy") & " " & Format(Time, "(HH:mm:ss)") & ": " & Str
  Log_Size = LOF(1)
 Close #1
 If Log_Size > 3145728 Then
  If Dir(App.Path & "\Logs", vbDirectory) = vbNullString Then MkDir App.Path & "\Logs"
  FileCopy App.Path & "\log.log", App.Path & "\log" & Format(Date, "dd.mm.yyyy") & "_" & Format(Time, "HH:mm:ss") & ".log"
  Kill App.Path & "\log.log"
 End If
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
 ShellExecute S_Main.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub
