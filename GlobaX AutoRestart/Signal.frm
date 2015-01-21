VERSION 5.00
Begin VB.Form Signal 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'Нет
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   2415
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
      Caption         =   "Ваш баланс менее 10 руб."
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
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Signal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim y As Integer
Dim t As Integer

Private Sub Form_Load()
 Label1.Caption = "Ваш баланс менее " & Trim(Min_Balans) & " руб."
 SetForegroundWindow Me.hwnd
 Me.Move Screen.Width - Me.Width, Screen.Height
 Me.Visible = True
End Sub

Private Sub Label1_Click()
 Unload Me
End Sub

Private Sub Timer_Timer()
 If y >= 450 + Me.Height Then
  t = t + 1
  If t >= 500 Then
   Unload Me
  End If
 Else
  y = y + (y / 50) + 1
  Me.Move Screen.Width - Me.Width, Screen.Height - y
 End If
End Sub
