VERSION 5.00
Begin VB.Form Signal 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'Нет
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Сейчас у Вас есть возможность играть!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Signal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Y As Integer
Dim t As Integer

Private Sub Form_Load()
 If Var_Bilet = 1 Then Label1.Caption = "Сейчас у Вас есть возможность играть!"
 If Var_Bilet = 2 Then Label1.Caption = "На продажу выставлены билеты!"
 SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
 Me.Move Screen.Width - Me.Width, Screen.Height
 Me.Visible = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then File_Run "http://www.freeloto.ru/auth.php?enter=ok&login=" & frmMain.TextLogin & "&pass_b64=" & AsciiToBase64(frmMain.TextPass)
 frmMain.Timer_Tray.Enabled = False
 Tray_Modify frmMain.Picture_An(0).Picture
 Unload Me
End Sub

Private Sub Timer_Timer()
 If Y >= 450 + Me.Height Then
  t = t + 1
  If t >= 500 Then
   Unload Me
  End If
 Else
  Y = Y + (Y / 50) + 1
  Me.Move Screen.Width - Me.Width, Screen.Height - Y
 End If
End Sub
