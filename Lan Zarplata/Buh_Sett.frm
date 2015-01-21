VERSION 5.00
Begin VB.Form Sett_Buh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "Buh_Sett.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Com_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   1980
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Com_Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   420
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text_T 
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text_T 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text_T 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text_T 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text_T 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3720
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label6 
      Caption         =   "час."
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Число рабочих часов:"
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
      TabIndex        =   15
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3720
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Рабочие часы"
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
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Время обеда"
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
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "мин."
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "мин."
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label26 
      Caption         =   "час."
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "час."
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label24 
      Caption         =   "Конец:"
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
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "Начало:"
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
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Sett_Buh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_Cancel_Click()
 Unload Me
End Sub

Private Sub Com_Ok_Click()
 Com_Ok.Enabled = False
 Com_Cancel.Enabled = False
 Dim Temp() As String
 Dim a As String
 a = SendCom("Put_Buh_Sett|" & Text_T(0).Text & "|" & Text_T(1).Text & "|" & Text_T(2).Text & "|" & Text_T(3).Text & "|" & Text_T(4).Text)
 If a = "Error" Then
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Else
  Unload Me
 End If
 Com_Ok.Enabled = True
 Com_Cancel.Enabled = True
End Sub

Private Sub Form_Load()
 Dim Temp() As String
 Dim a As String
 a = SendCom("Get_Buh_Sett")
 If a = "Error" Then
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Else
  Temp() = Split(a, "|")
  Text_T(0).Text = Temp(0)
  Text_T(1).Text = Temp(1)
  Text_T(2).Text = Temp(2)
  Text_T(3).Text = Temp(3)
  Text_T(4).Text = Temp(4)
 End If
End Sub

Private Sub Text_T_GotFocus(Index As Integer)
 Text_T(Index).SelStart = 0
 Text_T(Index).SelLength = 2
End Sub

Private Sub Text_T_KeyPress(Index As Integer, KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 Or KeyAscii <> 44) Then KeyAscii = 0
End Sub

Private Sub Text_T_Validate(Index As Integer, Cancel As Boolean)
 If Index = 1 Or Index = 3 Then
  If Val(Text_T(Index)) = 0 Or (Val(Text_T(Index)) < 0 Or Val(Text_T(Index)) > 59) Then Text_T(Index).Text = "00"
 Else
  If Val(Text_T(Index)) = 0 Or (Val(Text_T(Index)) < 0 Or Val(Text_T(Index)) > 23) Then Text_T(Index).Text = "00"
 End If
End Sub
