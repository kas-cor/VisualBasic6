VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Main_Client 
   Caption         =   "CallBack"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "Main_Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Общие настройки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Нет
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   3495
         TabIndex        =   14
         Top             =   240
         Width           =   3495
         Begin VB.TextBox Text_Manager 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Text            =   "Manager"
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Com_SavSet 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox Check_AutoRun 
            Caption         =   "Запускать при старте системы"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "Имя менеджера:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   3255
         End
      End
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
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Нет
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3495
         TabIndex        =   10
         Top             =   240
         Width           =   3495
         Begin VB.TextBox Text_Host 
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Text            =   "127.0.0.1"
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox Text_Port 
            Height          =   285
            Left            =   960
            TabIndex        =   1
            Text            =   "1000"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Com_SavNet 
            Caption         =   "Сохранить"
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Сервер:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Порт:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   615
         End
      End
   End
   Begin VB.Timer Timer_Connect 
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton Com_Connect 
      Caption         =   "Подключится"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   3735
   End
   Begin VB.PictureBox PicHook 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   1000
   End
   Begin VB.Label Label_Status 
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Статус"
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
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu_HiSh 
         Caption         =   "Показать/скрыть"
      End
      Begin VB.Menu Menu_About 
         Caption         =   "О программе..."
      End
      Begin VB.Menu Menu_s1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exit 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "Main_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Flag_Connect As Boolean
Dim Gl_Flag_Err As Boolean
Dim Tmr_Con As Long
Dim Min_Flag As Boolean

Private Function Connect(H As String, P As Long) As Boolean
 On Error Resume Next
 Dim Tmr As Long
 Dim i As Long
 Dim aP As String
 Dim Flag_Err As Boolean
 Flag_Connect = False
 Gl_Flag_Err = False
 Winsock.Close
 Add_Status "Подключение к " & H & ":" & P & "..."
 Winsock.Connect H, P
 Tmr = Timer()
 Do
  DoEvents
  If Timer() - Tmr > 60 Then
   Add_Status "Таймаут " & H & ":" & P
   Flag_Err = True
  End If
 Loop Until Flag_Connect Or Flag_Err Or Gl_Flag_Err
 If Flag_Connect Then
  aP = SendCom("Get_Free_Port")
  Winsock.Close
  If aP <> "Error" Then
   Add_Status "Подключение к " & H & ":" & CLng(aP) & "..."
   Winsock.Connect H, CLng(aP)
   Tmr = Timer()
   Do
    DoEvents
    If Timer() - Tmr > 60 Then
     Add_Status "Таймаут " & H & ":" & CLng(aP)
     Flag_Err = True
    End If
   Loop Until Flag_Connect Or Flag_Err Or Gl_Flag_Err
  Else
   Connect = False
  End If
  Connect = Flag_Connect
 End If
End Function

Private Sub Com_Connect_Click()
 If Connect(Host, CLng(Port)) Then
  Timer_Connect.Enabled = False
  Com_Connect.Enabled = False
  Com_Connect.Caption = "Подключено"
  Add_Status "Подключился"
 End If
End Sub

Private Sub Com_SavNet_Click()
 Host = Text_Host.Text
 Port = Text_Port.Text
 SaveSetting "CallBack", "Setting", "Host", Host
 SaveSetting "CallBack", "Setting", "Port", Port
End Sub

Private Sub Com_SavSet_Click()
 Manager = Text_Manager.Text
 SaveSetting "CallBack", "Setting", "Manager", Manager
End Sub

Private Sub Form_Load()
 Randomize Timer()
 Tmr_Con = Int(Rnd(1) * 60)
 Port = GetSetting("CallBack", "Setting", "Port", 1000)
 Host = GetSetting("CallBack", "Setting", "Host", "127.0.0.1")
 Manager = GetSetting("CallBack", "Setting", "Manager", "Manager")
 Text_Host.Text = Host
 Text_Port.Text = Port
 Text_Manager.Text = Manager
 Check_AutoRun.Value = GetSetting("CallBack", "Setting", "AutoRun", 0)
 Tray_Add Me.Icon, Me.Caption, PicHook.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting "CallBack", "Setting", "AutoRun", Check_AutoRun.Value
 If Check_AutoRun.Value = 1 Then
  RegWrite "Software\Microsoft\Windows\CurrentVersion\Run", "CallBack", App.Path & "\" & App.EXEName & ".exe"
 Else
  RegDelete "Software\Microsoft\Windows\CurrentVersion\Run", "CallBack"
 End If
 Tray_Del PicHook.hwnd
 Set Main_Client = Nothing
 End
End Sub

Private Sub Timer_Connect_Timer()
 If Not Min_Flag And Check_AutoRun.Value = 1 Then Me.WindowState = 1: Min_Flag = True
 If Tmr_Con < 120 Then
  Com_Connect.Caption = "Подключится [" & (120 - Tmr_Con) & "]"
  Tmr_Con = Tmr_Con + 1
 Else
  Tmr_Con = Int(Rnd(1) * 60)
  Call Com_Connect_Click
 End If
End Sub

Private Sub Winsock_Close()
 Winsock.Close
 Flag_Connect = False
 Tmr_Con = Int(Rnd(1) * 60)
 Timer_Connect.Enabled = True
 Com_Connect.Enabled = True
 Add_Status "Отключился"
End Sub

Private Sub Winsock_Connect()
 Flag_Connect = True
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
 Dim d As String
 Winsock.GetData d
 Add_Status "Получено " & Len(d) & " из " & bytesTotal & " байт..."
 Data = Data & d
 If Mid$(d, 1, 7) = "Message" Then
  If Op_WinSignal <> True And Op_WinInfo <> True Then
   Message = d
   Signal.Show vbModeless, Me
  End If
  Data = vbNullString
 End If
 If Mid$(d, 1, 12) = "Close_Signal" Then
  Call CloseSignal(d)
  Data = vbNullString
 End If
 If Mid$(d, 1, 6) = "Update" Then
  Call UpDateMe(d)
  Data = vbNullString
 End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Flag_Connect = False
 Tmr_Con = Int(Rnd(1) * 60)
 Timer_Connect.Enabled = True
 Com_Connect.Enabled = True
 Add_Status "#" & Number & " - " & Description
 Gl_Flag_Err = True
 Winsock.Close
End Sub

Private Sub Winsock_SendComplete()
 Add_Status "Отправлено"
End Sub

Private Sub Winsock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
 Add_Status "Отправлено " & Int((bytesRemaining + bytesSent) - bytesSent) & " из " & (bytesRemaining + bytesSent) & " байт..."
End Sub

Private Sub Menu_About_Click()
 About.Show vbModeless, Me
End Sub

Private Sub Menu_Exit_Click()
 Unload Me
 End
End Sub

Private Sub Menu_HiSh_Click()
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
 Dim Msg As Long
 Msg = X / Screen.TwipsPerPixelX
 If Msg = WM_LBUTTONDBLCLK Then
  Call Menu_HiSh_Click
 ElseIf Msg = WM_RBUTTONUP Then
  SetForegroundWindow Me.hwnd
  PopupMenu Menu, , , , Menu_HiSh
  PostMessage Me.hwnd, WM_NULL, 0, 0
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
  Call Menu_HiSh_Click
  Cancel = True
 Else
  Unload Me
 End If
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If Me.WindowState = 1 Then
  Me.Visible = False
 Else
  Me.Enabled = False
  Me.Height = 5205
  Me.Width = 4095
  Me.Enabled = True
 End If
End Sub

Private Sub CloseSignal(d As String)
 Dim Temp() As String
 Temp() = Split(d, "#;#")
 Cl_WinSignal_Men = Temp(1)
 Cl_WinSignal = True
End Sub

Private Sub UpDateMe(d As String)
 Dim Temp() As String
 Temp() = Split(d, "#;#")
 FileCopy Mid$(Temp(1), 1, Len(Temp(1)) - 2), App.Path & "\update.exe"
 Open App.Path & "\start.cmd" For Output As #1
  Print #1, "@echo off"
  Print #1, "start " & App.Path & "\update.cmd"
  Print #1, "del .\start.cmd"
 Close #1
 Open App.Path & "\update.cmd" For Output As #1
  Print #1, "@echo off"
  Print #1, "echo - - - - - - - - - - - - - - - -"
  Print #1, "echo - Begining update CallBack... -"
  Print #1, "echo - - - - - - - - - - - - - - - -"
  Print #1, "echo Delete old version..."
  Print #1, "del " & App.EXEName & ".exe"
  Print #1, "echo Copy new version..."
  Print #1, "copy update.exe " & App.EXEName & ".exe"
  Print #1, "echo Runing new virsion..."
  Print #1, "start " & App.EXEName & ".exe"
  Print #1, "echo Delete temp files..."
  Print #1, "echo Update complete."
  Print #1, "echo - - - - - - - - - - - - - - - -"
  Print #1, "echo -  !!! Close this window !!!  -"
  Print #1, "echo - - - - - - - - - - - - - - - -"
  Print #1, "del update.*"
 Close #1
 Shell App.Path & "\start.cmd", vbHide
 Unload Me
 Set Main_Client = Nothing
 End
End Sub
