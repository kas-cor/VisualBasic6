VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Запуск аудиорекламы"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5295
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_Play 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2520
   End
   Begin VB.CommandButton Command_Save 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text_Min 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "10"
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer Timer_Reklama 
      Interval        =   1000
      Left            =   600
      Top             =   2520
   End
   Begin VB.Timer Timer_Volume 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "localhost"
      RemotePort      =   806
   End
   Begin VB.CommandButton cmdDown 
      Height          =   495
      Left            =   3480
      Picture         =   "frmClient.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Volume Down"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Height          =   495
      Left            =   3000
      Picture         =   "frmClient.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Volume Up"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPause 
      Height          =   495
      Left            =   1080
      Picture         =   "frmClient.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdShuffle 
      Height          =   495
      Left            =   2520
      Picture         =   "frmClient.frx":1110
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Shuffle"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   495
      Left            =   120
      Picture         =   "frmClient.frx":1552
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Previous Track"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      Height          =   495
      Left            =   2040
      Picture         =   "frmClient.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Stop"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Height          =   495
      Left            =   1560
      Picture         =   "frmClient.frx":1DD6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Next Track"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   495
      Left            =   600
      Picture         =   "frmClient.frx":2218
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label_Status 
      BorderStyle     =   1  'Фиксировано один
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
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "мин."
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Реклама через каждые:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Dim T_Interval As Long
Dim T_Reklama As Long
Dim T_Volume As Long
Dim T_Play As Long

Private Sub cmdNext_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "Next"
End Sub

Private Sub cmdPause_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "Pause"
End Sub

Private Sub cmdPlay_Click()
    Winsock.Connect
    waitConnect
    Winsock.SendData "Play"
End Sub

Private Sub waitConnect()
    While Winsock.State <> sckConnected
      DoEvents
    Wend
End Sub

Private Sub cmdPrevious_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "Previous"
End Sub

Private Sub cmdShuffle_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "Shuffle"
End Sub

Private Sub cmdStop_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "Stop"
End Sub

Private Sub cmdUp_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "+2"
End Sub

Private Sub cmdDown_Click()
  Winsock.Connect
  waitConnect
  Winsock.SendData "-2"
End Sub

Private Sub Command_Save_Click()
 T_Interval = Val(Text_Min.Text)
 SaveSetting "AIMP", "Control", "Min", T_Interval
End Sub

Private Sub Form_Load()
 T_Interval = GetSetting("AIMP", "Control", "Min", 10)
 Text_Min.Text = T_Interval
End Sub

Private Sub Form_Unload(Cancel As Integer)
 T_Interval = Val(Text_Min.Text)
 SaveSetting "AIMP", "Control", "Min", T_Interval
 Set frmClient = Nothing
 End
End Sub


Private Sub Timer_Reklama_Timer()
 T_Reklama = T_Reklama + 1
 Label_Status.Caption = "До начала проигрыша рекламы осталось: " & T_Interval * 60 - T_Reklama & " сек."
 If T_Reklama >= T_Interval * 60 Then
  T_Reklama = 0
  Timer_Reklama.Enabled = False
  Timer_Volume.Enabled = True
 End If
End Sub

Private Sub Timer_Volume_Timer()
 T_Volume = T_Volume + 1
 If T_Volume <= 10 Then
  Label_Status.Caption = "Увеличение громкости до 100%"
  Call cmdUp_Click
 ElseIf T_Volume > 10 And T_Volume < 20 Then
  Label_Status.Caption = "Уменьшение громкости до 10%"
  Call cmdDown_Click
 ElseIf T_Volume = 20 Then
  Timer_Volume.Enabled = False
  Timer_Play.Enabled = True
 ElseIf T_Volume > 20 And T_Volume < 30 Then
  Label_Status.Caption = "Увеличение громкости до 100%"
  Call cmdUp_Click
 ElseIf T_Volume = 30 Then
  T_Volume = 0
  Timer_Volume.Enabled = False
  Timer_Reklama.Enabled = True
 End If
End Sub

Private Sub Timer_Play_Timer()
 T_Play = T_Play + 1
 Label_Status.Caption = "Проигрываем рекламу, осталось: " & 60 - T_Play & " сек."
 If T_Play >= 60 Then
  MP3_Stop "MyAlias"
  T_Play = 0
  Timer_Play.Enabled = False
  Timer_Volume.Enabled = True
 ElseIf T_Play = 1 Then
  MP3_Play App.Path & "\sound.mp3", "MyAlias"
 End If
End Sub

Private Sub Winsock_SendComplete()
  Winsock.Close
End Sub

Private Function MP3_Play(ByVal sFile As String, ByVal sAlias As String) As Boolean
 
 Dim bResult As Boolean
 Dim sBuffer As String
 Dim lResult As Long
 
 sBuffer = Space$(255)
 lResult = GetShortPathName(sFile, sBuffer, Len(sBuffer))
 
 If lResult <> 0 Then
  sFile = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  lResult = mciSendString("open " & sFile & " type MPEGVideo alias " & sAlias, 0, 0, 0)
  If lResult = 0 Then
   If mciSendString("play " & sAlias & " from 0", 0, 0, 0) = 0 Then
    bResult = True
   End If
  End If
 End If
 MP3_Play = bResult
 
End Function

Public Sub MP3_Stop(ByVal sAlias As String)
 mciSendString "stop " & sAlias, 0, 0, 0
 mciSendString "close " & sAlias, 0, 0, 0
End Sub

