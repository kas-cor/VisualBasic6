VERSION 5.00
Begin VB.Form Signal 
   BorderStyle     =   0  'Нет
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "Signal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer_Color 
      Interval        =   50
      Left            =   0
      Top             =   960
   End
   Begin VB.PictureBox Pic_Men 
      BorderStyle     =   0  'Нет
      Height          =   2775
      Left            =   3480
      ScaleHeight     =   2775
      ScaleWidth      =   2895
      TabIndex        =   7
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton Com_KnopClose 
         Caption         =   "Закрыть"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label_Men 
         Alignment       =   2  'Центровка
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   9
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Центровка
         Caption         =   "Звонок принял менеджер"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.TextBox Text_Comment 
      BackColor       =   &H8000000F&
      Height          =   855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Timer Timer_Close 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton Com_Cancel 
      Caption         =   "Отклонить"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Com_Ok 
      Caption         =   "Принять"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Тема обращения:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label_Org 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label_Region 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Регион:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Организация:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3015
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

Dim Y As Long
Dim Tmr_Cl As Long
Dim Id As Long

Private Sub Com_Cancel_Click()
 Unload Me
End Sub

Private Sub Com_KnopClose_Click()
 Tmr_Cl = 1001
End Sub

Private Sub Com_Ok_Click()
 Dim a As String
 a = SendCom("Manager#;#" & Manager & "#;#" & Id)
 If a <> "Error" Then
  Full_Info.Show vbModeless, Main_Client
  Unload Me
 Else
  Com_Ok.Enabled = False
 End If
End Sub

Private Sub Form_Load()
 Randomize Timer()
 Dim Temp() As String
 Op_WinSignal = True
 Temp() = Split(Message, "#;#")
 Id = CLng(Temp(1))
 Label_Org.Caption = Temp(2)
 Label_Region.Caption = Temp(6)
 Text_Comment.Text = Temp(7)
 SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
 Me.Move Screen.Width - Me.Width, Screen.Height
 Me.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Op_WinSignal = False
 Set Signal = Nothing
End Sub

Private Sub Timer_Close_Timer()
 If Cl_WinSignal Then
  If Tmr_Cl = 0 Then
   Pic_Men.Move 240, 240
   Label_Men.Caption = Cl_WinSignal_Men
  End If
  Tmr_Cl = Tmr_Cl + 1
  If Tmr_Cl > 1000 Then
   Cl_WinSignal = False
   Unload Me
  End If
 End If
End Sub

Private Sub Timer_Color_Timer()
 Dim r1 As Long
 Dim r2 As Long
 r1 = Int(Rnd(1) * 60)
 r2 = Int(Rnd(1) * 60)
 Shape1.Move 120 - r1, 120 - r1, 3135 + r2, 3015 + r2
 Shape1.BorderColor = RGB(Int(Rnd(1) * 255), Int(Rnd(1) * 255), Int(Rnd(1) * 255))
End Sub

Private Sub Timer1_Timer()
 If Y >= 450 + Me.Height Then
  Timer1.Enabled = False
 Else
  Y = Y + (400 - Y / 10)
  Me.Move Screen.Width - Me.Width, Screen.Height - Y
 End If
End Sub
