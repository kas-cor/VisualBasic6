VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Информер FreeLoto.ru"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check_Prod 
      Caption         =   "Проверять выставленные на продажу билеты"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.PictureBox Picture_An 
      Height          =   375
      Index           =   1
      Left            =   2040
      Picture         =   "main.frx":0ABA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture_An 
      Height          =   375
      Index           =   0
      Left            =   1560
      Picture         =   "main.frx":0C04
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer_Tray 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   3120
   End
   Begin VB.CheckBox Check_Run 
      Caption         =   "Запускать программу при старте Windows"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.PictureBox PicHook 
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   1296
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
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3120
   End
   Begin VB.TextBox TextPass 
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
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TextLogin 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   4080
      Y1              =   1320
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Status2 
      Alignment       =   2  'Центровка
      Caption         =   "www.FreeLoto.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "main.frx":0D4E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Status 
      Alignment       =   2  'Центровка
      Caption         =   "До проверки осталось 10 мин. 0 сек."
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
      TabIndex        =   6
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Пароль:"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Логин:"
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
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu Menu 
      Caption         =   "Меню"
      Visible         =   0   'False
      Begin VB.Menu Menu_GoSite 
         Caption         =   "www.FreeLoto.ru"
      End
      Begin VB.Menu Menu_s2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_ShHi 
         Caption         =   "Показать/Скрыть"
      End
      Begin VB.Menu Menu_s1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer
Dim Flag As Boolean
Dim An_Tray As Integer

Private Sub Form_Load()
 If App.PrevInstance Then End
 WebBrowser.Navigate "About:<b>Empty</b>"
 TextLogin.Text = GetSetting("InfoFreeLoto", "Setting", "Login", "")
 TextPass.Text = Pass_DeCrypt(GetSetting("InfoFreeLoto", "Setting", "Pass", ""))
 Check_Run.Value = GetSetting("InfoFreeLoto", "Setting", "Run", 0)
 Check_Prod.Value = GetSetting("InfoFreeLoto", "Setting", "Prod", 0)
 TextPass.PasswordChar = Chr$(149)
 Tray_Add Picture_An(0).Picture, Me.Caption, PicHook.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status2.FontUnderline = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
  Call Menu_ShHi_Click
  Cancel = True
 End If
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 1 Then
  Me.Visible = False
 Else
  Me.Enabled = False
  Me.Height = 2445
  Me.Width = 4335
  Me.Enabled = True
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting "InfoFreeLoto", "Setting", "Login", TextLogin.Text
 SaveSetting "InfoFreeLoto", "Setting", "Pass", Pass_EnCrypt(TextPass.Text)
 SaveSetting "InfoFreeLoto", "Setting", "Run", Check_Run.Value
 SaveSetting "InfoFreeLoto", "Setting", "Prod", Check_Prod.Value
 If Check_Run.Value = 1 Then
  RegWrite "Software\Microsoft\Windows\CurrentVersion\Run", "InfoFreeLoto", App.Path + "\" + App.EXEName + ".exe"
 Else
  RegDelete "Software\Microsoft\Windows\CurrentVersion\Run", "InfoFreeLoto"
 End If
 Tray_Del PicHook.hwnd
 End
End Sub

Private Sub Menu_Exit_Click()
 Unload Me
End Sub

Private Sub Menu_GoSite_Click()
 File_Run "http://www.freeloto.ru/auth.php?enter=ok&login=" & frmMain.TextLogin & "&pass_b64=" & AsciiToBase64(frmMain.TextPass)
 Timer_Tray.Enabled = False
 Tray_Modify Picture_An(0).Picture
 Me.WindowState = 1
End Sub

Private Sub Menu_ShHi_Click()
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
 Dim Msg As Integer
 Msg = X / Screen.TwipsPerPixelX
 If Msg = WM_LBUTTONDBLCLK Then
  Call Menu_GoSite_Click
 ElseIf Msg = WM_LBUTTONDOWN Then
  Call Menu_ShHi_Click
 ElseIf Msg = WM_RBUTTONUP Then
  SetForegroundWindow Me.hwnd
  PopupMenu Menu, , , , Menu_GoSite
  PostMessage Me.hwnd, WM_NULL, 0, 0
 End If
End Sub

Private Sub Status2_Click()
 File_Run "http://www.freeloto.ru/auth.php?enter=ok&login=" & frmMain.TextLogin & "&pass_b64=" & AsciiToBase64(frmMain.TextPass)
 Timer_Tray.Enabled = False
 Tray_Modify Picture_An(0).Picture
End Sub

Private Sub Status2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status2.FontUnderline = True
End Sub

Private Sub Timer_Tray_Timer()
 If An_Tray = 2 Then An_Tray = 0
 Tray_Modify Picture_An(An_Tray).Picture
 An_Tray = An_Tray + 1
End Sub

Private Sub Timer1_Timer()
 Dim Tmr As Double
 Dim Txt As String
 If Flag = False Then Me.WindowState = 1: Flag = True
 If t > 60 * 10 - 1 Then
  t = 0
  Randomize Timer
  WebBrowser.Navigate "http://freeloto.ru/info.php?login=" & TextLogin.Text & "&pass=" & TextPass.Text & "&rnd=" & Int(Rnd(1) * 100000)
  Status.Caption = "Проверка..."
  Tmr = Timer
  Do
   DoEvents
   If Timer - Tmr > 60 Then Exit Do
  Loop Until WebBrowser.ReadyState = READYSTATE_COMPLETE
  Txt = WebBrowser.Document.body.innertext
  If Val(Txt) > 0 Then
   Var_Bilet = 1
   Signal.Show vbModeless, Me
   Status2.Caption = "У Вас есть возможность играть!"
   Tray_Modify_Tip "Информер FreeLoto.ru" & vbCrLf & "У Вас есть возможность играть!"
   Timer_Tray.Enabled = True
  ElseIf Txt = "salle" And Check_Prod.Value = 1 Then
   Var_Bilet = 2
   Signal.Show vbModeless, Me
   Status2.Caption = "На продажу выставлены билеты!"
   Tray_Modify_Tip "Информер FreeLoto.ru" & vbCrLf & "На продажу выставлены билеты!"
   Timer_Tray.Enabled = True
  ElseIf Txt = "error" Then
   Status2.Caption = "Неверные логин или пароль!"
   Tray_Modify_Tip "Информер FreeLoto.ru" & vbCrLf & "Неверные логин или пароль!"
  Else
   Status2.Caption = "www.FreeLoto.ru"
   Tray_Modify_Tip "Информер FreeLoto.ru"
   Timer_Tray.Enabled = False
   Tray_Modify Picture_An(0).Picture
  End If
 End If
 Status.Caption = "До проверки осталось " & 9 - Int(t / 60) & " мин. " & 59 - (t - (Int(t / 60) * 60)) & " сек."
 t = t + 1
End Sub
