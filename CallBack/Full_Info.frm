VERSION 5.00
Begin VB.Form Full_Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Информация о клиенте"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Full_Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Com_EMail 
      Caption         =   "Написать"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Com_Print 
      Caption         =   "На печать"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3540
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton Com_Close 
      Cancel          =   -1  'True
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   6
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   5
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   5
      Top             =   3600
      Width           =   6495
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox Text_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Комментарий:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   6495
   End
   Begin VB.Label Label5 
      Caption         =   "Регион:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Телефон:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Контактное лицо:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Организация:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Full_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_Close_Click()
 Unload Me
End Sub

Private Sub Com_EMail_Click()
 Call File_Run("mailto:" & Text_Info(3).Text & "?subject=Обратный звонок&body=Здравствуйте, " & Text_Info(1).Text & " !")
End Sub

Private Sub Com_Print_Click()
 Dim i As Long
 Open App.Path & "\client.htm" For Output As #1
  Print #1, "<html><head>"
  Print #1, "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1251"">"
  Print #1, "<meta http-equiv=""Content - Language"" content=""ru"">" & vbCrLf
  Print #1, "<title>Информация о клиенте - Печать</title>"
  Print #1, "</head>"
  Print #1, "<body>"
  Print #1, "<div align=""center"">"
  Print #1, "<table border=""0"" cellpadding=""0"" style=""border-collapse: collapse"" width=""100%"">"
  Print #1, "<tr>"
  Print #1, "<td align=""center"" width=""100%"">"
  Print #1, "<div align=""center"">"
  Print #1, "<table border=""1"" cellpadding=""3"" style=""border-collapse: collapse"" width=""50%"" cellspacing=""3"">"
  Print #1, "<tr>"
  Print #1, "<td align=""Right"" valign=""top"" width=""30""><b>Организация:</b></td>"
  Print #1, "<td align=""Left"" valign=""top"" width=""70"">&nbsp;" & Text_Info(0) & "</td>"
  Print #1, "</tr>"
  Print #1, "<tr>"
  Print #1, "<td align=""right"" valign=""top"" width=""30%""><b>Контактное лицо:</b></td>"
  Print #1, "<td align=""left"" valign=""top"" width=""70%"">&nbsp;" & Text_Info(1) & "</td>"
  Print #1, "</tr>"
  Print #1, "<tr>"
  Print #1, "<td align=""right"" valign=""top"" width=""30%""><b>Телефон:</b></td>"
  Print #1, "<td align=""left"" valign=""top"" width=""70%"">&nbsp;" & Text_Info(2) & "</td>"
  Print #1, "</tr>"
  Print #1, "<tr>"
  Print #1, "<td align=""right"" valign=""top"" width=""30%""><b>e-mail:</b></td>"
  Print #1, "<td align=""left"" valign=""top"" width=""70%"">&nbsp;<a href=""mailto:" & Text_Info(3) & """>" & Text_Info(3) & "</a></td>"
  Print #1, "</tr>"
  Print #1, "<tr>"
  Print #1, "<td align=""right"" valign=""top"" width=""30%""><b>Регион:</b></td>"
  Print #1, "<td align=""left"" valign=""top"" width=""70%"">&nbsp;" & Text_Info(4) & "</td>"
  Print #1, "</tr>"
  Print #1, "<tr>"
  Print #1, "<td align=""right"" valign=""top"" width=""30%""><b>Комментарий:</b></td>"
  Print #1, "<td align=""left"" valign=""top"" width=""70%""><pre>" & Text_Info(5) & "</pre></td>"
  Print #1, "</tr>"
  Print #1, "</table>"
  Print #1, "</div>"
  Print #1, "</td>"
  Print #1, "</tr>"
  Print #1, "</table>"
  Print #1, "</div>"
  Print #1, "</body></html>"
 Close #1
 Call File_Run(App.Path & "\client.htm")
End Sub

Private Sub Form_Load()
 Dim i As Long
 Dim Temp() As String
 Op_WinInfo = True
 Temp() = Split(Message, "#;#")
 For i = 0 To 5
  Text_Info(i).Text = IIf(Temp(i + 2) <> vbNullString, Temp(i + 2), "Не указанно")
 Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Op_WinInfo = False
 Set Full_Info = Nothing
End Sub

Private Sub Text_Info_GotFocus(Index As Integer)
 Text_Info(Index).SelStart = 0
 Text_Info(Index).SelLength = Len(Text_Info(Index))
End Sub
