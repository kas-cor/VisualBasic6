VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form S_Main 
   Caption         =   "Сервер программы Зарплата"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   Icon            =   "Mail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   -120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Привязать вниз
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   9869
            Text            =   "Статус"
            TextSave        =   "Статус"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_Tray 
      Interval        =   1
      Left            =   -120
      Top             =   960
   End
   Begin VB.CheckBox Check_AutoRun 
      Caption         =   "Запускать программу при старте системы"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   5655
   End
   Begin VB.PictureBox PicHook 
      Height          =   375
      Left            =   -120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer_Del 
      Interval        =   1000
      Left            =   -120
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Привилегированные пользователи"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   5655
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Нет
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5415
         TabIndex        =   13
         Top             =   240
         Width           =   5415
         Begin VB.CommandButton Save_Priv 
            Caption         =   "Сохранить"
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox Adm_Login 
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Text            =   "Admin"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Adm_Pass 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3480
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox B_Login 
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Text            =   "Buh"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox B_Pass 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3480
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Пароль:"
            Height          =   255
            Left            =   2640
            TabIndex        =   19
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Бухгалтер"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   5175
         End
         Begin VB.Label Label3 
            Caption         =   "Администратор"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   5175
         End
         Begin VB.Label Label5 
            Caption         =   "Логин:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Пароль:"
            Height          =   255
            Left            =   2640
            TabIndex        =   15
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Логин:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Настройки сети"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Нет
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5415
         TabIndex        =   11
         Top             =   240
         Width           =   5415
         Begin VB.CommandButton Com_SavePort 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   3120
            TabIndex        =   1
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox Port 
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Text            =   "1000"
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Порт:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu Menu_HiSh 
         Caption         =   "Показать/Скрыть"
      End
      Begin VB.Menu Menu_About 
         Caption         =   "О программе..."
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "S_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Flag_Connect(0 To 1000) As Boolean ' Флаг коннекта
Dim data(0 To 1000) As String '          Получаемые данные
Dim Timer_Day As Long '                Таймер 10 мин.

Private Sub Com_SavePort_Click()
 Ports = CLng(Port.Text)
 Call Save_Setting
 MsgBox "Для вступление изменений в силу, перезапустите программу.", vbInformation, "Предупреждение"
End Sub

Private Sub Form_Load()
 If App.PrevInstance Then End
 On Error Resume Next
 Adm_Pass.PasswordChar = Chr$(149)
 B_Pass.PasswordChar = Chr$(149)
 Dim i As Long
 Crypt_Key = Get_Crypt_Key
 ' Пользователи
 Open App.Path & "\Users.dat" For Input As #1
  Input #1, Users_Max
  For i = 0 To Users_Max - 1
   Input #1, Users_Login(i), Users_Pass(i), Users_ModeStav(i), Users_ZP(i), Users_Stav(i), Users_Edit_Time(i), Users_Enter_Obed(i), Users_Path(i)
   Users_Pass(i) = Pass_DeCrypt(Users_Pass(i))
  Next
 Close #1
 ' Не рабочие дни
 Open App.Path & "\DontWork.dat" For Input As #1
  Input #1, DontWork_Max
  For i = 0 To DontWork_Max - 1
   Input #1, DontWork_Date(i)
  Next
 Close #1
 ' Настройки
 Admin_Login = "Admin"
 Admin_Pass = "123"
 Buh_Login = "Buh"
 Buh_Pass = "123"
 Ports = 1000
 StartObed_Hr = 13
 StartObed_Min = 0
 EndObed_Hr = 14
 EndObed_Min = 0
 Work_Hr = 8
 Open App.Path & "\Setting.dat" For Input As #1
  Input #1, Admin_Login, Admin_Pass, Buh_Login, Buh_Pass, Ports, StartObed_Hr, StartObed_Min, EndObed_Hr, Work_Hr
  Admin_Pass = Pass_DeCrypt(Admin_Pass)
  Buh_Pass = Pass_DeCrypt(Buh_Pass)
 Close #1
 Check_AutoRun.Value = GetSetting("Zarplata", "Setting", "AutoRun", 0)
 Adm_Login.Text = Admin_Login
 Adm_Pass.Text = Admin_Pass
 B_Login.Text = Buh_Login
 B_Pass.Text = Buh_Pass
 Port.Text = Ports
 For i = 0 To 20
  Winsock(i).LocalPort = Ports + i
  Winsock(i).Listen
 Next
 Call Mk_Dirs
 Timer_Day = 1000
 Tray_Add Me.Icon, Me.Caption, PicHook.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
  Call Menu_HiSh_Click
  Cancel = True
 End If
End Sub

Private Sub Menu_About_Click()
 About.Show vbModeless, Me
End Sub

Private Sub Menu_Exit_Click()
 Unload Me
 End
End Sub

Private Sub Menu_HiSh_Click()
 If Me.WindowState = 1 Then
  Me.WindowState = 0
  Me.Visible = True
  SetForegroundWindow Me.hwnd
 Else
  Me.WindowState = 1
  Me.Visible = False
 End If
End Sub

Private Sub PicHook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Msg As Long
 Msg = X / Screen.TwipsPerPixelX
 If Msg = WM_LBUTTONDBLCLK Then
  Call Menu_HiSh_Click
 ElseIf Msg = WM_RBUTTONUP Then
  SetForegroundWindow Me.hwnd
  PopupMenu Menu, , , , Menu_HiSh
  PostMessage Me.hwnd, WM_NULL, 0, 0
 End If
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If Me.WindowState = 1 Then
  Me.Visible = False
 Else
  Me.Enabled = False
  Me.Height = 4965
  Me.Width = 6015
  Me.Enabled = True
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Save_Setting
 Call Save_Users
 Call Save_DontWork
 SaveSetting "Zarplata", "Setting", "AutoRun", Check_AutoRun.Value
 If Check_AutoRun.Value = 1 Then
  RegWrite "Software\Microsoft\Windows\CurrentVersion\Run", "ZarplataLAN", App.Path + "\" + App.EXEName + ".exe"
 Else
  RegDelete "Software\Microsoft\Windows\CurrentVersion\Run", "ZarplataLAN"
 End If
 Tray_Del PicHook.hwnd
 End
End Sub

Private Sub Save_Priv_Click()
 Admin_Login = Adm_Login.Text
 Admin_Pass = Adm_Pass.Text
 Buh_Login = B_Login.Text
 Buh_Pass = B_Pass.Text
 Call Save_Setting
End Sub

Private Sub StatusBar_PanelDblClick(ByVal Panel As ComctlLib.Panel)
 Form_Status.Visible = True
End Sub

Private Sub Timer_Del_Timer()
 On Error Resume Next
 Dim d As Long
 Dim i As Long
 If Timer_Day > 600 Then
  Timer_Day = 0
  d = GetSetting("Zarplata", "Setting", "Last_Day", 0)
  If d <> Day(Date) Then
   For i = 0 To Users_Max - 1
    Kill App.Path & "\base\" & Users_Path(i) & "\temp.dat"
   Next
  End If
  SaveSetting "Zarplata", "Setting", "Last_Day", Day(Date)
 End If
 Timer_Day = Timer_Day + 1
End Sub

Private Sub Timer_Tray_Timer()
 Me.WindowState = 1
 Timer_Tray.Enabled = False
End Sub

Private Sub Winsock_Close(Index As Integer)
 Flag_Connect(Index) = False
 Winsock(Index).Close
 If Index = 0 Then
  Winsock(0).LocalPort = Ports
  Winsock(0).Listen
 End If
 SID_Status(Index) = False
 Add_Status Index, "Отключился " & Winsock(Index).RemoteHostIP & ":" & Winsock(Index).RemotePort
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
 Winsock(Index).Close
 Winsock(Index).Accept requestID
 Flag_Connect(Index) = True
 SID_Status(Index) = True
 Add_Status Index, "Подключился " & Winsock(Index).RemoteHostIP & ":" & Winsock(Index).RemotePort
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim d As String
 Winsock(Index).GetData d
 Add_Status Index, "Получено " & Len(d) & " из " & bytesTotal & " байт..."
 data(Index) = data(Index) & ZP_DeCode(d)
 If InStr(1, data(Index), "." & vbCrLf) <> 0 Then
  d = Mid$(data(Index), 1, Len(data(Index)) - 3)
  
  ' Авторизация
  If d = "Get_Free_Ports" Then ' Список свободных портов
   Winsock(Index).SendData ZP_EnCode(GetFreePorts() & "." & vbCrLf)
  End If
  If d = "Get_Users_List" Then ' Список пользователей
   Winsock(Index).SendData ZP_EnCode(UsersList() & "." & vbCrLf)
  End If
  If Mid$(d, 1, 6) = "Login " Then ' Запрос пароля пользователя
   Winsock(Index).SendData ZP_EnCode(UserPass(Mid$(d, 7)) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 8) = "SID for " Then ' Выдача сессии
   Winsock(Index).SendData ZP_EnCode(Get_SID(Index, Mid$(d, 9)) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 9) = "Get_Type " Then ' Выдача типа пользователя
   Winsock(Index).SendData ZP_EnCode(SID_Type(CLng(Mid$(d, 10))) & "." & vbCrLf)
  End If
  
  ' Работа с пользователем
  If Mid$(d, 1, 12) = "Get_Time_Now" Then ' Выдача текущего времени
   Winsock(Index).SendData ZP_EnCode(Unix_Time(Hour(Time()), Minute(Time()), Day(Date), Month(Date), Year(Date)) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 10) = "Get_Login " Then ' Выдача логина
   Winsock(Index).SendData ZP_EnCode(Users_Login(SID_LogNum(CLng(Mid$(d, 11)))) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 14) = "Data_For_User " Then ' Выдача данных пользователя
   Winsock(Index).SendData ZP_EnCode(Data_For_User(Mid$(d, 15)) & "." & vbCrLf)
  End If
  
  ' Работа с данными
  If Mid$(d, 1, 9) = "Get_Data " Then ' Вывод данных
   Winsock(Index).SendData ZP_EnCode(Get_Data(CLng(Mid$(d, 10))) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 6) = "Record" Then ' Запись данных
   Winsock(Index).SendData ZP_EnCode(Record(d) & "." & vbCrLf)
  End If
  
  ' Администрирование
  If Mid$(d, 1, 9) = "Edit_User" Then ' Изменение пользователя
   Winsock(Index).SendData ZP_EnCode(Edit_User(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 8) = "Add_User" Then ' Добавление пользователя
   Winsock(Index).SendData ZP_EnCode(Add_User(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 12) = "Delete_User " Then ' Удвление пользователя
   Winsock(Index).SendData ZP_EnCode(Delete_User(Mid$(d, 13)) & "." & vbCrLf)
  End If
  
  ' Бухгалтер
  If Mid$(d, 1, 9) = "Change_ZP" Then ' Начисление зарплаты
   Winsock(Index).SendData ZP_EnCode(Change_ZP(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 11) = "Change_Stav" Then ' Изменение ставки в час
   Winsock(Index).SendData ZP_EnCode(Change_Stav(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 6) = "Report" Then ' Создание отчета
   Winsock(Index).SendData ZP_EnCode(Report(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 16) = "Get_DontWorkDays" Then ' Вывод не рабочих дней
   Winsock(Index).SendData ZP_EnCode(Get_DontWorkDays() & "." & vbCrLf)
  End If
  If Mid$(d, 1, 16) = "Put_DontWorkDays" Then ' Вывод не рабочих дней
   Winsock(Index).SendData ZP_EnCode(Put_DontWorkDays(d) & "." & vbCrLf)
  End If
  If Mid$(d, 1, 12) = "Get_Buh_Sett" Then ' Вывод настроек бухгалтера
   Winsock(Index).SendData ZP_EnCode(Get_Buh_Sett() & "." & vbCrLf)
  End If
  If Mid$(d, 1, 12) = "Put_Buh_Sett" Then ' Запись настроек бухгалтера
   Winsock(Index).SendData ZP_EnCode(Put_Buh_Sett(d) & "." & vbCrLf)
  End If
  
  ' Отправка шаблонов
  If Mid$(d, 1, 12) = "Get_Shablon1" Then ' Шаблон простого отчета
   Winsock(Index).SendData ZP_EnCode(Get_Shablon1() & "." & vbCrLf)
  End If
  If Mid$(d, 1, 12) = "Get_Shablon2" Then ' Шаблон отчета по все работникам
   Winsock(Index).SendData ZP_EnCode(Get_Shablon2() & "." & vbCrLf)
  End If
  
  data(Index) = vbNullString
 End If
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Winsock(Index).Close
 Flag_Connect(Index) = False
 If Index = 0 Then
  Winsock(0).LocalPort = Ports
  Winsock(0).Listen
 End If
 SID_Status(Index) = False
 Add_Status Index, "#" & Number & " - " & Description
End Sub

Private Sub Winsock_SendComplete(Index As Integer)
 Add_Status Index, "Отправлено"
End Sub

Private Sub Winsock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
 Add_Status Index, "Отправлено " & Int((bytesRemaining + bytesSent) - bytesSent) & " из " & (bytesRemaining + bytesSent) & " байт..."
End Sub

Function GetFreePorts()
 ' Вывод списка свободных портов
 Dim i As Long
 If Winsock.Count > 1 Then
  GetFreePorts = -1
  For i = 1 To Winsock.Count - 1
   If Not Flag_Connect(i) Then
    GetFreePorts = Ports + i
    Exit For
   End If
  Next
  If GetFreePorts = -1 Then GetFreePorts = Ports + NewSock: i = Winsock.Count - 1
 Else
  GetFreePorts = Ports + NewSock: i = Winsock.Count - 1
 End If
 Winsock(i).LocalPort = GetFreePorts
 Winsock(i).Listen
End Function

Private Function NewSock() As Long
 Dim c As Long
 c = Winsock.Count
 Load Winsock(c)
 NewSock = c
End Function

Function UsersList() As String
 ' Вывод списка пользователей
 Dim i As Long
 Dim s As String
 For i = 0 To Users_Max - 1
  s = s & Users_Login(i) & vbCrLf
 Next
 s = s & Admin_Login & vbCrLf
 s = s & Buh_Login & vbCrLf
 UsersList = s
End Function

Function UserPass(Log As String) As String
 ' Вывод пароля пользователя
 Dim i As Long
 If Log = Admin_Login Then
  UserPass = Admin_Pass & "|" & String$(Rnd(1) * 50, "*")
  Exit Function
 End If
 If Log = Buh_Login Then
  UserPass = Buh_Pass & "|" & String$(Rnd(1) * 50, "*")
  Exit Function
 End If
 For i = 0 To Users_Max - 1
  If Log = Users_Login(i) Then
   UserPass = Users_Pass(i) & "|" & String$(Rnd(1) * 50, "*")
   Exit Function
  End If
 Next
 UserPass = "Error"
End Function

Function Get_SID(Index As Integer, Log As String) As Long
 ' Выдача сесии пользователю
 Dim i As Long
 Select Case Log
  Case Admin_Login ' Администратор
   SID_Login(Index) = Log
   SID_Type(Index) = 2
   Get_SID = Index
   Exit Function
  Case Buh_Login ' Бухгалтер
   SID_Login(Index) = Log
   SID_Type(Index) = 1
   Get_SID = Index
   Exit Function
 End Select
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Log Then ' Пользователь
   SID_Login(Index) = Log
   SID_LogNum(Index) = i
   SID_Type(Index) = 0
   Get_SID = Index
   Exit Function
  End If
 Next
 Get_SID = Index
End Function

Function Data_For_User(Log As String)
 ' Выдача данных пользователя
 Dim i As Long
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Log Then Exit For
 Next
 Data_For_User = IIf(Users_ModeStav(i), "1", "0") & "|" & Users_ZP(i) & "|" & Users_Stav(i) & "|" & IIf(Users_Edit_Time(i), "1", "0") & "|" & IIf(Users_Enter_Obed(i), "1", "0") & "|" & Users_Path(i)
End Function

Function Get_Data(s As Long) As String
 ' Запрос данных
 Dim i As Long
 Dim d(0 To 3) As Long
 For i = 0 To Users_Max - 1
  If Users_Login(i) = SID_Login(s) Then Exit For
 Next
 If Dir(App.Path & "\base\" & Users_Path(i) & "\temp.dat") = vbNullString Then
  d(0) = 0
  d(1) = 0
  d(2) = 0
  d(3) = 0
 Else
  Open App.Path & "\base\" & Users_Path(i) & "\temp.dat" For Input As #1
   Input #1, d(0), d(1), d(2), d(3)
  Close #1
 End If
 Get_Data = d(0) & "," & d(1) & "," & d(2) & "," & d(3)
End Function

Function Record(d As String) As String
 ' Запись данных (приход, обед, уход)
 Dim Temp() As String
 Dim Tmp As String
 Dim i As Long
 Dim s As Long ' Сессия
 Dim d1 As String
 Dim d2(0 To 3) As Long
 Temp() = Split(d, "|")
 s = CInt(Temp(1))
 d1 = Temp(2)
 Temp() = Split(d1, ",")
 d2(0) = Temp(0)
 d2(1) = Temp(1)
 d2(2) = Temp(2)
 d2(3) = Temp(3)
 For i = 0 To Users_Max - 1
  If Users_Login(i) = SID_Login(s) Then Exit For
 Next
 If Not Users_Enter_Obed(i) And d2(1) = 0 And d2(2) = 0 Then
  Tmp = Time_Unix(d2(0))
  d2(1) = Unix_Time(StartObed_Hr, StartObed_Min, CLng(Mid$(Tmp, 7, 2)), CLng(Mid$(Tmp, 10, 2)), CLng(Mid$(Tmp, 13, 4)))
  d2(2) = Unix_Time(EndObed_Hr, EndObed_Min, CLng(Mid$(Tmp, 7, 2)), CLng(Mid$(Tmp, 10, 2)), CLng(Mid$(Tmp, 13, 4)))
 End If
 Open App.Path & "\base\" & Users_Path(i) & "\temp.dat" For Output As #1
  Write #1, d2(0), d2(1), d2(2), d2(3)
 Close #1
 If d2(0) <> 0 And d2(1) <> 0 And d2(2) <> 0 And d2(3) <> 0 Then
  Open App.Path & "\base\" & Users_Path(i) & "\data.dat" For Append As #1
   Write #1, d2(0), d2(1), d2(2), d2(3)
  Close #1
 End If
 Record = "Ok"
End Function

Function Edit_User(d As String) As String
 ' Изменение данных пользователя
 Dim i As Long
 Dim Temp() As String
 Temp() = Split(d, "|")
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Temp(1) Then Exit For
 Next
 Users_Login(i) = Temp(2)
 Users_Pass(i) = Temp(3)
 Users_Edit_Time(i) = Temp(4)
 Users_Enter_Obed(i) = Temp(5)
 Users_Path(i) = Temp(6)
 Call Save_Users
 Call Mk_Dirs
 Edit_User = "Ok"
End Function

Function Add_User(d As String) As String
 ' Добавление нового пользователя
 Dim i As Long
 Dim Flag As Boolean
 Dim Temp() As String
 Temp() = Split(d, "|")
 For i = 0 To Users_Max - 1
  If Users_Path(i) = Temp(4) Then Flag = True
 Next
 If Flag Then
  Randomize Timer
  Temp(4) = Temp(4) & "_" & Trim(Int(Rnd(1) * 9999 + 10))
 End If
 Users_Login(Users_Max) = Temp(1)
 Users_Pass(Users_Max) = Temp(2)
 Users_Edit_Time(Users_Max) = Temp(3)
 Users_Enter_Obed(i) = Temp(4)
 Users_Path(Users_Max) = Temp(5)
 Users_Max = Users_Max + 1
 Call Save_Users
 Call Mk_Dirs
 Add_User = "Ok"
End Function

Function Delete_User(Log As String) As String
 ' Удаление пользователя
 Dim i As Long
 Dim t As Long
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Log Then Exit For
 Next
 For t = i To Users_Max - 1
  Users_Login(t) = Users_Login(t + 1)
  Users_Pass(t) = Users_Pass(t + 1)
  Users_ZP(t) = Users_ZP(t + 1)
  Users_Stav(t) = Users_Stav(t + 1)
  Users_Edit_Time(t) = Users_Edit_Time(t + 1)
  Users_Enter_Obed(t) = Users_Enter_Obed(t + 1)
  Users_Path(t) = Users_Path(t + 1)
 Next
 Users_Max = Users_Max - 1
 Call Save_Users
 Delete_User = "Ok"
End Function

Function Change_ZP(d As String) As String
 ' Начисление зарплаты
 Dim i As Long
 Dim Temp() As String
 Temp() = Split(d, "|")
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Temp(1) Then Exit For
 Next
 Users_ZP(i) = CLng(Temp(2))
 Call Save_Users
 Change_ZP = "Ok"
End Function

Function Change_Stav(d As String) As String
 ' Изменение ставки в час
 Dim i As Long
 Dim Temp() As String
 Temp() = Split(d, "|")
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Temp(1) Then Exit For
 Next
 Users_ModeStav(i) = IIf(CLng(Temp(2)) = 1, True, False)
 Users_ZP(i) = CLng(Temp(3))
 Users_Stav(i) = CLng(Temp(4))
 Call Save_Users
 Change_Stav = "Ok"
End Function

Function Report(d As String) As String
 ' Создание отчета
 Dim i As Long
 Dim s As Long
 Dim e As Long
 Dim a As String
 Dim t(0 To 3) As Long
 Dim Temp() As String
 Temp() = Split(d, "|")
 For i = 0 To Users_Max - 1
  If Users_Login(i) = Temp(1) Then Exit For
 Next
 s = CLng(Temp(2))
 e = CLng(Temp(3)) + 1440
 If Dir(App.Path & "\base\" & Users_Path(i) & "\data.dat", vbNormal) <> vbNullString Then
  Open App.Path & "\base\" & Users_Path(i) & "\data.dat" For Input As #1
   Do Until EOF(1)
    Input #1, t(0), t(1), t(2), t(3)
    If t(0) >= s And t(0) <= e Then a = a & t(0) & "," & t(1) & "," & t(2) & "," & t(3) & vbCrLf
   Loop
  Close #1
 End If
 If a <> vbNullString Then Report = a Else Report = "Ok"
End Function

Function Get_DontWorkDays() As String
 ' Вывод не рабочих дней
 Dim i As Long
 Dim a As String
 For i = 0 To DontWork_Max - 1
  a = a & Trim(CStr(DontWork_Date(i))) & vbCrLf
 Next
 Get_DontWorkDays = a
End Function

Function Put_DontWorkDays(d As String) As String
 ' Запись не рабочих дней
 Dim i As Long
 Dim Temp() As String
 Temp() = Split(d, vbCrLf)
 DontWork_Max = UBound(Temp()) - 1
 For i = 0 To DontWork_Max - 1
  DontWork_Date(i) = CLng(Temp(i + 1))
 Next
 Call Save_DontWork
 Put_DontWorkDays = "Ok"
End Function

Function Get_Buh_Sett() As String
 ' Вывод настроек бухгалтера
 Dim a As String
 a = StartObed_Hr & "|" & _
 StartObed_Min & "|" & _
 EndObed_Hr & "|" & _
 EndObed_Min & "|" & _
 Work_Hr
 Get_Buh_Sett = a
End Function

Function Put_Buh_Sett(d As String) As String
 ' Запись настроек бухгалтера
 Dim a As String
 Dim Temp() As String
 Dim i As Long
 Temp() = Split(d, "|")
 StartObed_Hr = CLng(Temp(1))
 StartObed_Min = CLng(Temp(2))
 EndObed_Hr = CLng(Temp(3))
 EndObed_Min = CLng(Temp(4))
 Work_Hr = CLng(Temp(5))
 Call Save_Setting
 Put_Buh_Sett = "Ok"
End Function

Function Get_Shablon1() As String
 ' Шаблон простого отчета
 On Error Resume Next
 Dim n As String
 Dim txt As String
 Open App.Path & "\templates\template1.html" For Input As #1
  Do
   Line Input #1, n
   txt = txt & n
  Loop Until EOF(1)
 Close #1
 If txt <> vbNullString Then Get_Shablon1 = txt Else Get_Shablon1 = "Error"
End Function

Function Get_Shablon2() As String
 ' Шаблон отчета по всем сотрудникам
 On Error Resume Next
 Dim n As String
 Dim txt As String
 Open App.Path & "\templates\template2.html" For Input As #1
  Do
   Line Input #1, n
   txt = txt & n
  Loop Until EOF(1)
 Close #1
 If txt <> vbNullString Then Get_Shablon2 = txt Else Get_Shablon2 = "Error"
End Function

Function Get_Crypt_Key()
 ' Ключ шифрования
 Dim i As Long
 Dim Temp As String
 Dim k As String
 If Dir(App.Path & "\key.txt", vbNormal) <> vbNullString Then
  Open App.Path & "\key.txt" For Input As #1
   Do
    Line Input #1, Temp
    k = k & Temp
   Loop Until EOF(1)
  Close #1
 Else
  Randomize Timer()
  For i = 0 To 1024
   k = k & Chr$(Rnd(1) * 255)
  Next
  Open App.Path & "\key.txt" For Output As #1
   Print #1, k
  Close #1
 End If
 Get_Crypt_Key = k
End Function

Private Sub Mk_Dirs()
 ' Создание необходимых дирикторий
 On Error Resume Next
 Dim i As Long
 Err.Clear
 If Dir(App.Path & "\base", vbDirectory) = vbNullString Then MkDir App.Path & "\base"
 If Err.Number <> 0 Then MsgBox "Не удалось создать папку " & App.Path & "\base, сделайте это самостоятельно!", vbCritical, "Ошибка"
 For i = 0 To Users_Max - 1
  Err.Clear
  If Dir(App.Path & "\base\" & Users_Path(i), vbDirectory) = vbNullString Then MkDir App.Path & "\base\" & Users_Path(i)
  If Err.Number <> 0 Then MsgBox "Не удалось создать папку " & App.Path & "\base\" & Users_Path(i) & ", сделайте это самостоятельно!", vbCritical, "Ошибка"
 Next
End Sub

Public Sub ClientCount()
 Dim i As Long
 Dim n As Long
 For i = 0 To Winsock.Count
  If Flag_Connect(i) Then n = n + 1
 Next
 Form_Status.Label_Clients.Caption = "Подключено клиентов " & n
End Sub
