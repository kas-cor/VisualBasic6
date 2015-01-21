VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GlobaX AutoRestart"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_StartGX 
      Interval        =   2000
      Left            =   4560
      Top             =   840
   End
   Begin VB.Timer Timer_Min 
      Interval        =   1000
      Left            =   600
      Top             =   5520
   End
   Begin VB.Frame Frame_LP 
      BorderStyle     =   0  'Нет
      Height          =   495
      Left            =   5760
      TabIndex        =   22
      Top             =   6480
      Width           =   4575
      Begin VB.TextBox Text_Pass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text_Login 
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Пароль:"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Логин:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton Cmd_About 
      Caption         =   "О программе..."
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'Нет
      Height          =   2655
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      Top             =   3720
      Width           =   5175
      Begin VB.Timer Timer_tarif 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   1440
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Изменить"
         Height          =   375
         Left            =   1680
         TabIndex        =   31
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Центровка
         Caption         =   "Перезапуск произойдет через: 0 сек."
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
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Центровка
         Caption         =   "Для заверщения необходим перезапуск GlobaX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label12 
         Caption         =   "Изменение тарифного плана:"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Ваш тарифный план:"
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
         TabIndex        =   28
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label_Chg_Tarif 
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   720
         Width           =   2895
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   5040
         Y1              =   500
         Y2              =   500
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'Нет
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   5175
      Begin VB.Timer Timer_Balans 
         Interval        =   1000
         Left            =   0
         Top             =   1800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Обновить"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label_Tarif 
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label_Status 
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label_Balans 
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label_Credit 
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Тарифный план:"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Статус блокировки:"
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
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Актуальный баланс:"
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
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Кредит:"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   5040
         Y1              =   500
         Y2              =   500
      End
      Begin VB.Label Label_UpDate 
         Alignment       =   2  'Центровка
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   3495
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'Нет
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5175
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   0
         TabIndex        =   35
         Top             =   600
         Width           =   5175
         Begin VB.TextBox Text_Interval 
            Height          =   285
            Left            =   1920
            TabIndex        =   40
            Text            =   "30"
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton Save 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   3240
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox Check_Restart 
            Caption         =   "Использовать автоматический перезапуск"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   4815
         End
         Begin VB.CheckBox Check_UpDateBalans 
            Caption         =   "Автоматически обновлять баланс каждые 10 минут"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   4815
         End
         Begin VB.CheckBox Check_AutoRun 
            Caption         =   "Запускать программу при старте Windows"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label3 
            Caption         =   "сек."
            Height          =   255
            Left            =   2640
            TabIndex        =   42
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Интервал проверки:"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Перезапустить"
         Height          =   495
         Left            =   3540
         TabIndex        =   34
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton GXStart 
         Caption         =   "Запустить"
         Height          =   495
         Left            =   60
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton GXStop 
         Caption         =   "Остановить"
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.PictureBox PicHook 
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer TimerGX 
         Interval        =   1000
         Left            =   4920
         Top             =   360
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      CausesValidation=   0   'False
      Height          =   3375
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
      Location        =   "http:///"
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":0442
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line7 
      X1              =   5400
      X2              =   120
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Label Status 
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
      TabIndex        =   7
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Line Line4 
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   3135
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   1800
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Caption         =   "Тариф STV"
      Height          =   255
      Index           =   2
      Left            =   3720
      MouseIcon       =   "Form1.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Caption         =   "Баланс STV"
      Height          =   255
      Index           =   1
      Left            =   1920
      MouseIcon       =   "Form1.frx":09D6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Caption         =   "GlobaX"
      Height          =   255
      Index           =   0
      Left            =   120
      MouseIcon       =   "Form1.frx":0B28
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Con_Setting 
         Caption         =   "Показать окно"
      End
      Begin VB.Menu Con_S4 
         Caption         =   "-"
      End
      Begin VB.Menu Con_Tarif 
         Caption         =   "Тарифный план"
         Begin VB.Menu Con_Chg_Trf 
            Caption         =   "Оптимальный 256"
            Index           =   0
         End
         Begin VB.Menu Con_Chg_Trf 
            Caption         =   "Оптимальный 512"
            Index           =   1
         End
         Begin VB.Menu Con_Chg_Trf 
            Caption         =   "Оптимальный 1024"
            Index           =   2
         End
         Begin VB.Menu Con_Chg_Trf 
            Caption         =   "Оптимальный 2048"
            Index           =   3
         End
         Begin VB.Menu Con_Chg_Trf 
            Caption         =   "Оптимальный 3072"
            Index           =   4
         End
      End
      Begin VB.Menu Con_Configs 
         Caption         =   "Конфигурации"
         Begin VB.Menu Con_Sel_Config 
            Caption         =   "Config №1"
            Index           =   0
         End
         Begin VB.Menu Con_Sel_Config 
            Caption         =   "Config №2"
            Index           =   1
         End
         Begin VB.Menu Con_Sel_Config 
            Caption         =   "Config №3"
            Index           =   2
         End
      End
      Begin VB.Menu Con_S1 
         Caption         =   "-"
      End
      Begin VB.Menu Con_Start 
         Caption         =   "Запустить GX"
      End
      Begin VB.Menu Con_Stop 
         Caption         =   "Остановить GX"
      End
      Begin VB.Menu Con_Restart 
         Caption         =   "Перезапуск GX"
      End
      Begin VB.Menu Con_S2 
         Caption         =   "-"
      End
      Begin VB.Menu Con_Exit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AutoRestart As Boolean
Dim Port As String
Dim Interval As Integer
Dim Sel_Conf As Integer
Dim t As Integer
Dim STV_Balans As String
Dim STV_Credit As String
Dim STV_Status As String
Dim STV_Tarif As String
Dim STV_DataUpDate As String
Dim Status_Cabinet As Boolean
Dim Status_Tarif As Boolean
Dim T_Tarif As Integer
Dim T_Balans As Integer
Dim Balans_Signal As Boolean
Dim Temp_S_GX As String
Dim Log_Size As Double
Dim Process_Id As Integer
Dim AutoRun As Boolean

Private Sub Main_Initialize()
 Call InitCommonControls
End Sub

Private Sub Check_Restart_Click()
 If Check_Restart.Value = 1 Then
  Label2.Enabled = True
  Text_Interval.Enabled = True
  Label3.Enabled = True
  Save.Enabled = True
 ElseIf Check_Restart.Value = 0 Then
  Label2.Enabled = False
  Text_Interval.Enabled = False
  Label3.Enabled = False
  Save.Enabled = False
 End If
End Sub

Private Sub Cmd_About_Click()
 About.Show vbModal, Me
End Sub

Private Sub Command1_Click()
 On Error Resume Next
 Dim t As Double
 Dim n As String
 Dim N_Bal As Double
 Command1.Enabled = False
 Me.Visible = True
 WebBrowser.Visible = True
 DoEvents
 Status_Cabinet = False
 WebBrowser.Navigate "https://cabinet.stv.su/?rnd=" & Int(Rnd(1) * 1000000000)
 Label_UpDate.Caption = "Переход на сайт..."
 t = Timer
 Do
  DoEvents
  If Timer - t > 30 Then Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
 Label_UpDate.Caption = "Вход в кабинет..."
 WebBrowser.Document.Forms(0).Item("p1").Value = Text_Login.Text
 WebBrowser.Document.Forms(0).Item("p2").Value = Text_Pass.Text
 WebBrowser.Document.Forms(0).Submit
 t = Timer
 Do
  DoEvents
  If Timer - t > 30 Then Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_INTERACTIVE
 Label_UpDate.Caption = "Получение данных..."
 n = WebBrowser.Document.body.innertext
 Call Get_Data(n)
 STV_DataUpDate = Format(Date, "dd.mm.yyyy") & " " & Format(Time, "(HH:mm:ss)")
 N_Bal = Val(STV_Balans) + Val(STV_Credit)
 If Min_Balans = 300 Then
  If N_Bal < 300 And N_Bal > 200 Then
   Signal.Show
   Min_Balans = 200
  End If
 ElseIf Min_Balans = 200 Then
  If N_Bal < 200 And N_Bal > 100 Then
   Signal.Show
   Min_Balans = 100
  End If
 ElseIf Min_Balans = 100 Then
  If N_Bal < 100 And N_Bal > 50 Then
   Signal.Show
   Min_Balans = 50
  End If
 ElseIf Min_Balans = 50 Then
  If N_Bal < 50 And N_Bal > 10 Then
   Signal.Show
   Min_Balans = 10
  End If
 ElseIf Min_Balans = 10 Then
  If N_Bal < 10 Then
   Signal.Show
   Min_Balans = 300
  End If
 End If
 Call UpDate
 WebBrowser.Visible = False
 If Me.WindowState = 1 Then Me.Visible = False
 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
 On Error Resume Next
 Dim Trf(0 To 4) As Integer
 Dim t As Double
 Dim n As String
 Command2.Enabled = False
 Trf(0) = 256
 Trf(1) = 512
 Trf(2) = 1024
 Trf(3) = 2048
 Trf(4) = 3072
 Status_Cabinet = False
 Status_Tarif = False
 WebBrowser.Visible = True
 WebBrowser.Navigate "https://cabinet.stv.su/?rnd=" & Int(Rnd(1) * 1000000000)
 Label_Chg_Tarif.Caption = "Изменяется..."
 t = Timer
 Do
  DoEvents
  If Timer - t > 30 Then Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
 WebBrowser.Document.Forms(0).Item("p1").Value = Text_Login.Text
 WebBrowser.Document.Forms(0).Item("p2").Value = Text_Pass.Text
 WebBrowser.Document.Forms(0).Submit
 t = Timer
 Do
  DoEvents
  If Timer - t > 30 Then Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_INTERACTIVE
 WebBrowser.Navigate "https://cabinet.stv.su/insat/gate.php?mod=plan&mode=chplan&m=ch&planid=" & Trf(Combo1.ListIndex)
 t = Timer
 Do
  DoEvents
  If Timer - t > 30 Then Exit Do
 Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
 n = WebBrowser.Document.body.innertext
 Call Get_Data(n)
 Timer_tarif.Enabled = True
 Label13.Visible = True
 Label14.Visible = True
 STV_Tarif = "Оптимальный " & Trf(Combo1.ListIndex)
 Call UpDate
 WebBrowser.Visible = False
 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
 Call Con_Restart_Click
End Sub

Private Sub Con_Chg_Trf_Click(Index As Integer)
 Combo1.ListIndex = Index
 Command2_Click
End Sub

Private Sub Con_Exit_Click()
 Unload Me
 End
End Sub

Private Sub Con_Restart_Click()
 Call WaitMe(0)
 Call Kill_GX
 If Log_Size > 3000000 Then Kill App.Path & "\client.log"
 Shell App.Path & "\globax_daemon.exe", vbHide
 Call WaitMe(1)
 Status.Caption = "Перезапущен"
End Sub

Private Sub Con_Sel_Config_Click(Index As Integer)
 Dim t As Double
 ChDir App.Path
 Call Kill_GX
 FileCopy App.Path & "\Config" & Index + 1 & ".txt", App.Path & "\globax.conf"
 t = Timer
 Do
  If Timer - t > 2 Then Exit Do
  DoEvents
 Loop
 Shell App.Path & "\globax_daemon.exe", vbHide
 Sel_Conf = Index
 Call UpDate
 Status.Caption = "Config заменен"
 DoEvents
End Sub

Private Sub Con_Setting_Click()
 If Me.WindowState = 1 Then
  Me.WindowState = 0
  Me.Visible = True
  SetForegroundWindow Me.hwnd
 Else
  Me.WindowState = 1
  Me.Visible = False
 End If
End Sub

Private Sub Con_Start_Click()
 Call GXStart_Click
End Sub

Private Sub Con_Stop_Click()
 Call GXStop_Click
End Sub

Private Sub Form_Load()
 Dim N_Bal As Double
 If App.PrevInstance Then
  End
 End If
 ChDir App.Path
 Me.Caption = "GlobaX AutoRestart " & App.Major & "." & App.Minor & "." & App.Revision
 WebBrowser.Silent = True
 WebBrowser.Navigate "About:blank"
 AutoRestart = GetSetting("GXRestart", "Setting", "AutoRestart", False)
 Port = GetSetting("GXRestart", "Setting", "Port", "3128")
 Interval = GetSetting("GXRestart", "Setting", "Interval", 30)
 Sel_Conf = GetSetting("GXRestart", "Setting", "Config", 0)
 Text_Login.Text = GetSetting("GXRestart", "Setting", "Login", "")
 Text_Pass.Text = GetSetting("GXRestart", "Setting", "Pass", "")
 STV_Balans = GetSetting("GXRestart", "Setting", "Balans", "")
 STV_Credit = GetSetting("GXRestart", "Setting", "Credit", "")
 STV_Status = GetSetting("GXRestart", "Setting", "Status", "")
 STV_Tarif = GetSetting("GXRestart", "Setting", "Tarif", "")
 STV_DataUpDate = GetSetting("GXRestart", "Setting", "DateUpDate", "")
 Check_UpDateBalans.Value = GetSetting("GXRestart", "Setting", "AutoUpDateBalans", 0)
 Temp_S_GX = GetSetting("GXRestart", "Setting", "Temp_S", "")
 AutoRun = GetSetting("GXRestart", "Setting", "AutoRun", True)
 Check_AutoRun.Value = IIf(AutoRun, 1, 0)
 Check_Restart.Value = IIf(AutoRestart, 1, 0)
 Text_Interval.Text = Interval
 Tray_Add Image1(0).Picture, Me.Caption, PicHook.hwnd
 Combo1.Clear
 Combo1.AddItem "Оптимальный 256"
 Combo1.AddItem "Оптимальный 512"
 Combo1.AddItem "Оптимальный 1024"
 Combo1.AddItem "Оптимальный 2048"
 Combo1.AddItem "Оптимальный 3072"
 Combo1.ListIndex = 0
 Call Label4_Click(0)
 Call UpDate
 N_Bal = Val(STV_Balans) + Val(STV_Credit)
 If N_Bal > 0 And N_Bal < 10 Then Min_Balans = 300
 If N_Bal > 10 And N_Bal < 50 Then Min_Balans = 10
 If N_Bal > 50 And N_Bal < 100 Then Min_Balans = 50
 If N_Bal > 100 And N_Bal < 200 Then Min_Balans = 100
 If N_Bal > 200 And N_Bal < 300 Then Min_Balans = 200
 If N_Bal > 300 Then Min_Balans = 300
 If Min_Balans = 0 Then Min_Balans = 300
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
  Me.WindowState = 1
  Cancel = True
 End If
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 1 Then
  Me.Visible = False
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 SaveSetting "GXRestart", "Setting", "Login", Text_Login.Text
 SaveSetting "GXRestart", "Setting", "Pass", Text_Pass.Text
 SaveSetting "GXRestart", "Setting", "Balans", STV_Balans
 SaveSetting "GXRestart", "Setting", "Credit", STV_Credit
 SaveSetting "GXRestart", "Setting", "Status", STV_Status
 SaveSetting "GXRestart", "Setting", "Tarif", STV_Tarif
 SaveSetting "GXRestart", "Setting", "DateUpDate", STV_DataUpDate
 SaveSetting "GXRestart", "Setting", "AutoUpDateBalans", Check_UpDateBalans.Value
 SaveSetting "GXRestart", "Setting", "Temp_S", Temp_S_GX
 SaveSetting "GXRestart", "Setting", "AutoRun", IIf(Check_AutoRun.Value = 1, True, False)
 Call Save_Click
 If AutoRun Then
   RegWrite "Software\Microsoft\Windows\CurrentVersion\Run", "GlobaX AutoRestart", App.Path + "\" + App.EXEName + ".exe"
 Else
   RegDelete "Software\Microsoft\Windows\CurrentVersion\Run", "GlobaX AutoRestart"
 End If
 Call GXStop_Click
 If Me.WindowState = 1 Then
  Me.WindowState = 0
  Me.Visible = True
 End If
 Tray_Del PicHook.hwnd
End Sub

Private Sub GXStart_Click()
 If Not GX_Present Then
  Call WaitMe(0)
  Shell App.Path & "\globax_daemon.exe", vbHide
  Call WaitMe(1)
 End If
 Status.Caption = "Запущен"
 DoEvents
End Sub

Private Sub GXStop_Click()
 Call WaitMe(0)
 Call Kill_GX
 Status.Caption = "Остановлен"
 Call WaitMe(1)
 DoEvents
End Sub

Private Sub Label4_Click(Index As Integer)
 Dim i As Integer
 Frame_LP.Visible = False
 Frame_LP.Move 240, 480
 For i = 0 To 2
  Label4(i).FontBold = False
  Frame(i).Visible = False
  Frame(i).Move 120, 480, 5175, 2655
 Next
 Label4(Index).FontBold = True
 Frame(Index).Visible = True
 If Index = 1 Or Index = 2 Then Frame_LP.Visible = True
End Sub

Private Sub Save_Click()
 AutoRestart = IIf(Check_Restart.Value = 1, True, False)
 Interval = Val(Text_Interval.Text)
 SaveSetting "GXRestart", "Setting", "AutoRestart", AutoRestart
 SaveSetting "GXRestart", "Setting", "Interval", Interval
 SaveSetting "GXRestart", "Setting", "Config", Sel_Conf
 Call UpDate
End Sub

Private Sub PicHook_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim Msg As Integer
 Msg = x / Screen.TwipsPerPixelX
 If Msg = WM_LBUTTONDBLCLK Then
  If Me.WindowState = 1 Then
   Me.WindowState = 0
   Me.Visible = True
   SetForegroundWindow Me.hwnd
  Else
   Me.WindowState = 1
   Me.Visible = False
  End If
 ElseIf Msg = WM_RBUTTONUP Then
  If Me.WindowState = 0 Then
   Con_Setting.Caption = "Скрыть окно"
  ElseIf Me.WindowState = 1 Then
   Con_Setting.Caption = "Показать окно"
  End If
  SetForegroundWindow Me.hwnd
  PopupMenu Menu, , , , Con_Setting
  PostMessage Me.hwnd, WM_NULL, 0, 0
 End If
End Sub

Private Sub Timer_Balans_Timer()
 T_Balans = T_Balans + 1
 If T_Balans >= 600 Then
  T_Balans = 0
  If Check_UpDateBalans.Value = 1 Then
   Call Command1_Click
  End If
 End If
End Sub

Private Sub Timer_Min_Timer()
 If AutoRun Then
  Me.WindowState = 1
  Me.Visible = False
 End If
 Timer_Min.Enabled = False
End Sub

Private Sub Timer_StartGX_Timer()
 Call GXStart_Click
 Timer_StartGX.Enabled = False
End Sub

Private Sub Timer_Tarif_Timer()
 T_Tarif = T_Tarif + 1
 Label14.Caption = "Перезапуск произойдет через: " & 180 - T_Tarif & " сек."
 If T_Tarif >= 180 Then
  Call Con_Restart_Click
  T_Tarif = 0
  Timer_tarif.Enabled = False
  Label13.Visible = False
  Label14.Visible = False
 End If
End Sub

Private Sub TimerGX_Timer()
 On Error Resume Next
 Dim i As Integer
 Dim s(5) As String
 If t >= Interval Then
  Status.Caption = "Проверка на зависание GlobaX..."
  DoEvents
  ChDir App.Path
  Open "client.log" For Input As #1
   Log_Size = LOF(1)
   Do
    DoEvents
    Line Input #1, s(0)
    For i = 5 To 1 Step -1
     s(i) = s(i - 1)
     DoEvents
    Next
   Loop While Not EOF(1)
  Close
  For i = 1 To 5
   If Mid$(s(i), 22) = "Close session, seems no data/activity at specified timeout" Then
    If s(i) <> Temp_S_GX Then
     Status.Caption = "GlobaX завис, перезапускаю..."
     DoEvents
     Call Con_Restart_Click
     Temp_S_GX = s(i)
     Exit For
    End If
   End If
  Next
  t = 0
 Else
  Status.Caption = "До проверки " & Interval - t & " сек..."
  DoEvents
  t = t + 1
 End If
End Sub

Private Sub UpDate()
 Dim i As Integer
 If AutoRestart Then TimerGX.Enabled = True Else TimerGX.Enabled = False
 For i = 0 To 2
  Con_Sel_Config(i).Checked = False
 Next
 Con_Sel_Config(Sel_Conf).Checked = True
 Label_Balans.Caption = STV_Balans
 Label_Credit.Caption = STV_Credit
 Label_Status.Caption = STV_Status
 Label_Tarif.Caption = STV_Tarif
 Label_UpDate.Caption = "Информация обновлена: " & vbCrLf & STV_DataUpDate
 Label_Chg_Tarif.Caption = STV_Tarif
 For i = 0 To 4
  Con_Chg_Trf(i).Checked = False
 Next
 If Right(STV_Tarif, 3) = "256" Then Con_Chg_Trf(0).Checked = True
 If Right(STV_Tarif, 3) = "512" Then Con_Chg_Trf(1).Checked = True
 If Right(STV_Tarif, 4) = "1024" Then Con_Chg_Trf(2).Checked = True
 If Right(STV_Tarif, 4) = "2048" Then Con_Chg_Trf(3).Checked = True
 If Right(STV_Tarif, 4) = "3072" Then Con_Chg_Trf(4).Checked = True
 If Check_UpDateBalans.Value = 1 Then
  Call Tray_Modify_Tip(Me.Caption & vbCrLf & _
  "Баланс: " & STV_Balans & vbCrLf & _
  "Кредит: " & STV_Credit)
  '"Тариф: " & STV_Tarif & vbCrLf & _
  '"Статус: " & STV_Status)
 End If
 DoEvents
End Sub

Private Sub Get_Data(Dat As String)
 Dim Temp As Integer
 Temp = InStr(1, Dat, "Актуальный Баланс:")
 If Temp <> 0 Then STV_Balans = Trim(Mid(Mid(Dat, Temp, InStr(Temp, Dat, vbCrLf) - Temp), 19))
 Temp = InStr(1, Dat, "Кредит:")
 If Temp <> 0 Then STV_Credit = Trim(Mid(Mid(Dat, Temp, InStr(Temp, Dat, vbCrLf) - Temp), 8))
 Temp = InStr(1, Dat, "Статус блокировки:")
 If Temp <> 0 Then STV_Status = Trim(Mid(Mid(Dat, Temp, InStr(Temp, Dat, vbCrLf) - Temp), 19))
 Temp = InStr(1, Dat, "Тарифный план:")
 If Temp <> 0 Then STV_Tarif = Trim(Mid(Mid(Dat, Temp, InStr(Temp, Dat, vbCrLf) - Temp), 15))
 Call UpDate
 If InStr(1, Dat, "Добро пожаловать в Личный Кабинет") <> 0 Then Status_Cabinet = True
 If InStr(1, Dat, "Ваш тарифный план") <> 0 Then Status_Tarif = True
 If InStr(1, Dat, "Введен неправильный логин или пароль") <> 0 Then
  Label_UpDate.Caption = "Не верный логин или пароль"
  Label_Chg_Tarif.Caption = "Не верный логин или пароль"
 End If
End Sub

Sub WaitMe(p As Integer)
 If p = 0 Then
  GXStart.Enabled = False
  GXStop.Enabled = False
  Command3.Enabled = False
 Else
  GXStart.Enabled = True
  GXStop.Enabled = True
  Command3.Enabled = True
 End If
End Sub

Sub Kill_GX()
 Dim tmp As Double
 If GX_Present Then
  Shell "taskkill /IM globax_daemon.exe /F", vbHide
  tmp = Timer
  Do
   DoEvents
   If Timer - tmp > 3 Then Exit Do
   If Not GX_Present Then Exit Do
  Loop
 End If
End Sub
