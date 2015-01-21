VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Main_Server 
   Caption         =   "CallBack - Server"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "Main_Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Com_UpdateClients 
      Caption         =   "Обновить клиенты"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Com_CheckSite 
      Caption         =   "Проверить заявки"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   3975
      Left            =   4440
      TabIndex        =   15
      Top             =   120
      Width           =   2775
      ExtentX         =   4895
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "Общие настройки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Нет
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   10
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox Check_AutoRun 
            Caption         =   "Запускать при старте системы"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   3255
         End
         Begin VB.CommandButton Com_SavSet 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text_Key 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Text            =   "<введите ключ шифрования>"
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox Text_Url 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Text            =   "http://host.ru/callback/callback.php"
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Ключ шифрования:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Опрашиваемый сайт:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   3255
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
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Нет
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3495
         TabIndex        =   7
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton Com_SavNet 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   2040
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox Text_Port 
            Height          =   285
            Left            =   720
            TabIndex        =   0
            Text            =   "1000"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Порт:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   615
         End
      End
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1000
   End
   Begin VB.Timer Timer_Check 
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox PicHook 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label_Status 
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Статус"
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
      TabIndex        =   13
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu_HiSh 
         Caption         =   "Показать/скрыть"
      End
      Begin VB.Menu Menu_About 
         Caption         =   "О программе..."
      End
      Begin VB.Menu Menu_s1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "Main_Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Flag_Connect(0 To 1000) As Boolean ' Флаг коннекта
Dim data(0 To 1000) As String '          Данные
Dim Tmr_Check As Long '                  Таймер проверки
Dim Calls_Max As Long '                  Всего звонков
Dim Calls_Id(0 To 1000) As Long '        Список идентификаторов
Dim Calls_Str(0 To 1000) As String '     Cписок звонков
Dim Min_Flag As Boolean '                Минимизинг флаг
Dim Last_Id As Boolean '                 Последний идентификатов звонка
Dim Tmr_Discon As Long '                 Таймер дисконекта

Private Sub Com_CheckSite_Click()
 Tmr_Check = 60 * 5
End Sub

Private Sub Com_SavNet_Click()
 Port = Text_Port.Text
 Call Save_Setting
End Sub

Private Sub Com_SavSet_Click()
 Check_Url = Text_Url.Text
 Crypt_Key = Text_Key.Text
 Call Save_Setting
End Sub

Private Sub Com_UpdateClients_Click()
 Dim n As String
 Dim t As Long
 n = GetSetting("CallBack", "Setting", "Update", "\\Net\Документы\CallBack\callback_client.exe")
 n = InputBox("Введите сетевой путь до образца программы", "Обновление", n)
 SaveSetting "CallBack", "Setting", "Update", n
 For t = 0 To Winsock.Count
  If Flag_Connect(t) Then
   Winsock(t).SendData "Update#;#" & n & vbCrLf
  End If
 Next
 Add_Status 0, "Данна комманда на обновление клиентов"
End Sub

Private Sub Form_Load()
 On Error Resume Next
 WebBrowser.Navigate "About:<b>Empty</b>"
 Check_Url = "http://invent-prom.ru/callback/callback.php"
 Crypt_Key = "1234567890"
 Port = "1000"
 Open App.Path & "\Setting.dat" For Input As #1
  Input #1, Port, Check_Url, Crypt_Key
 Close #1
 Text_Port.Text = Port
 Text_Url.Text = Check_Url
 Text_Key.Text = Crypt_Key
 Check_AutoRun.Value = GetSetting("CallBack", "Setting", "AutoRun", 0)
 Winsock(0).LocalPort = Port
 Winsock(0).Listen
 Tray_Add Me.Icon, Me.Caption, PicHook.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Save_Setting
 SaveSetting "CallBack", "Setting", "AutoRun", Check_AutoRun.Value
 If Check_AutoRun.Value = 1 Then
  RegWrite "Software\Microsoft\Windows\CurrentVersion\Run", "CallBack", App.Path & "\" & App.EXEName & ".exe"
 Else
  RegDelete "Software\Microsoft\Windows\CurrentVersion\Run", "CallBack"
 End If
 Tray_Del PicHook.hwnd
 Set Main_Server = Nothing
 End
End Sub

Private Sub Label_Status_DblClick()
 Logi.Show vbModeless, Me
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
  Call Menu_HiSh_Click
  Cancel = True
 Else
  Unload Me
 End If
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If Me.WindowState = 1 Then
  Me.Visible = False
 Else
  Me.Enabled = False
  Me.Height = 5085
  Me.Width = 4095
  Me.Enabled = True
 End If
End Sub

Private Sub Timer_Check_Timer()
 If Not Min_Flag And Check_AutoRun.Value = 1 Then Me.WindowState = 1: Min_Flag = True
 If Tmr_Check < 60 * 5 Then
  Tmr_Check = Tmr_Check + 1
 Else
  Last_Id = False
  Tmr_Check = 0
  Call Rassilka(GetDataOnSite(Check_Url))
 End If
End Sub

Private Sub Winsock_Close(Index As Integer)
 Winsock(Index).Close
 Flag_Connect(Index) = False
 If Index = 0 Then
  Winsock(0).LocalPort = Port
  Winsock(0).Listen
 End If
 Add_Status Index, "Отключился " & Winsock(Index).RemoteHostIP & ":" & Winsock(Index).RemotePort
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
 Winsock(Index).Close
 Winsock(Index).Accept requestID
 Flag_Connect(Index) = True
 Add_Status Index, "Подключился " & Winsock(Index).RemoteHostIP & ":" & Winsock(Index).RemotePort
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim d As String
 Dim a As String
 Dim t As Long
 Winsock(Index).GetData d
 Add_Status Index, "Получено " & Len(d) & " из " & bytesTotal & " байт..."
 data(Index) = data(Index) & d
 If InStr(1, data(Index), "." & vbCrLf) <> 0 Then
  d = Mid$(data(Index), 1, Len(data(Index)) - 3)
  
  If d = "Get_Free_Port" Then ' Получение порта
   a = GetFreePort()
   Winsock(Index).SendData a & "." & vbCrLf
   Add_Status Index, "Присвоен порт " & a
  End If
  
  If Mid$(d, 1, 7) = "Manager" Then ' Присвоение звонка менеджеру
   a = PostManager(d)
   Winsock(Index).SendData a & "." & vbCrLf
   If a <> "Error" Then
    For t = 0 To Winsock.Count
     If Flag_Connect(t) And t <> Index Then
      Winsock(t).SendData "Close_Signal#;#" & a & vbCrLf
     End If
    Next
   End If
   Add_Status Index, "Звонок принял менеджер " & a
  End If
  
  data(Index) = vbNullString
 End If
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Winsock(Index).Close
 Flag_Connect(Index) = False
 If Index = 0 Then
  Winsock(0).LocalPort = Port
  Winsock(0).Listen
 End If
 Add_Status Index, "#" & Number & " - " & Description
End Sub

Private Sub Winsock_SendComplete(Index As Integer)
 Add_Status Index, "Отправлено"
End Sub

Private Sub Winsock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
 Add_Status Index, "Отправлено " & Int((bytesRemaining + bytesSent) - bytesSent) & " из " & (bytesRemaining + bytesSent) & " байт..."
End Sub

Private Function GetFreePort() As Long
 Dim i As Long
 If Winsock.Count > 1 Then
  GetFreePort = -1
  For i = 1 To Winsock.Count - 1
   If Not Flag_Connect(i) Then
    GetFreePort = Port + i
    Exit For
   End If
  Next
  If GetFreePort = -1 Then GetFreePort = Port + NewSock: i = Winsock.Count - 1
 Else
  GetFreePort = Port + NewSock: i = Winsock.Count - 1
 End If
 Winsock(i).LocalPort = GetFreePort
 Winsock(i).Listen
End Function

Private Function NewSock() As Long
 Dim c As Long
 c = Winsock.Count
 Load Winsock(c)
 NewSock = c
End Function

Private Function PostManager(d As String) As String
 If Last_Id Then PostManager = "Error": Exit Function
 Dim Flag As Boolean
 Dim Tmr As Long
 Dim Temp() As String
 Temp() = Split(d, "#;#")
 Randomize Timer()
 Add_Status 0, "Отправка данных на сайт..."
 WebBrowser.Navigate2 Check_Url & "?mark_call=" & Temp(2) & "&manager=" & Temp(1) & "&key=" & RC4_EnCode(Crypt_Key) & "&rnd=" & Rnd(1) * 9999999
 Tmr = Timer()
 Do
  DoEvents
  If Timer - Tmr > 60 Then Flag = True: Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
 Add_Status 0, "Готово"
 If Not Flag Then
  Last_Id = True
  PostManager = Temp(1)
 Else
  PostManager = "Error"
 End If
End Function

Private Function GetDataOnSite(Url As String) As String
 Dim Tmr As Long
 Dim n As String
 Dim i As Long
 Dim Flag As Boolean
 Randomize Timer()
 Add_Status 0, "Получение данных с сайта..."
 WebBrowser.Navigate2 Url & "?calls=list&rnd=" & Rnd(1) * 9999999
 Tmr = Timer()
 Do
  DoEvents
  If Timer - Tmr > 60 Then Flag = True: Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
 If Not Flag Then
  Add_Status 0, "Готово"
  n = WebBrowser.Document.body.innertext
  Flag = False
  For i = 1 To Len(n)
   If Not ((Asc(Mid$(n, i, 1)) > 47 And Asc(Mid$(n, i, 1)) < 58) Or (Asc(Mid$(n, i, 1)) > 64 And Asc(Mid$(n, i, 1)) < 71)) Then Flag = True: Exit For
  Next
  If Flag Then
   GetDataOnSite = "Empty"
  Else
   GetDataOnSite = RC4_DeCode(n)
  End If
 Else
  Add_Status 0, "Таймаут"
  GetDataOnSite = "Empty"
 End If
End Function

Private Sub Rassilka(txt As String)
 If txt = "Empty" Then Exit Sub
 Dim i As Long
 Dim t As Long
 Dim n As Long
 Dim TempLine() As String
 TempLine() = Split(txt, "{LineBreak}")
 Calls_Max = 0
 For i = 0 To UBound(TempLine()) - 1
  Calls_Str(Calls_Max) = Trim(TempLine(i))
  Calls_Max = Calls_Max + 1
 Next
 If Calls_Max <> 0 Then
  For t = 0 To Winsock.Count
   If Flag_Connect(t) Then
    Winsock(t).SendData "Message#;#" & Calls_Str(0) & vbCrLf
    n = n + 1
   End If
  Next
  Add_Status 0, "Разослоно сообщение на " & n & " менеджер(ов)"
 End If
End Sub

Public Sub ClientCount()
 Dim i As Long
 Dim n As Long
 For i = 0 To Winsock.Count
  If Flag_Connect(i) Then n = n + 1
 Next
 Logi.Label_Clients.Caption = "Подключено клиентов " & n
End Sub
