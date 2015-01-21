VERSION 5.00
Begin VB.Form Sett 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "Setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Com_Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Com_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
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
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
      Begin VB.TextBox TimeOut 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "60"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Host 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Port 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "1000"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "сек."
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Таймаут:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Адрес сервера:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Порт:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Sett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_Cancel_Click()
 Unload Me
End Sub

Private Sub Com_Ok_Click()
 Host = Trim(Host.Text)
 Ports = Val(Port.Text)
 TimeOut = Val(TimeOut.Text)
 SaveSetting "Zarplata", "Setting", "Host", Host
 SaveSetting "Zarplata", "Setting", "Ports", Ports
 SaveSetting "Zarplata", "Setting", "TimeOut", TimeOut
 Unload Me
End Sub

Private Sub Form_Load()
 Host.Text = Host
 Port.Text = Ports
 TimeOut.Text = TimeOut
End Sub
