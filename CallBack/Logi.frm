VERSION 5.00
Begin VB.Form Logi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Логи"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "Logi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Com_Close 
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.ListBox List_Log 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label_Clients 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   4815
   End
End
Attribute VB_Name = "Logi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_Close_Click()
 Unload Me
 Set Logi = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload Me
End Sub
