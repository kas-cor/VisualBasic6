VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "О программе"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "support@kas-cor.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1380
      MouseIcon       =   "About.frx":23D2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "http://kas-cor.ru/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1500
      MouseIcon       =   "About.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Центровка
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Made in Russia"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   480
      Left            =   1920
      Picture         =   "About.frx":2676
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "E-MAIL:"
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "WWW:"
      Height          =   255
      Left            =   900
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "KAS-cor"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Автор: kas-cor"
      Height          =   255
      Left            =   1260
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Label6.Caption = "CallBack " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set About = Nothing
End Sub

Private Sub Label7_Click(Index As Integer)
 If Index = 0 Then
  File_Run "http://kas-cor.ru/"
 ElseIf Index = 1 Then
  File_Run "mailto:support@kas-cor.ru"
 End If
End Sub
