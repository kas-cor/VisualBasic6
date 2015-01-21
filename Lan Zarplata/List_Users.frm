VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form List_Users 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Зарплата"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5175
   Icon            =   "List_Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Привязать вниз
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   9075
            Text            =   "Статус"
            TextSave        =   "Статус"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_Connect 
      Interval        =   500
      Left            =   120
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Com_Exit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Com_Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ListBox List 
      Enabled         =   0   'False
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Пользователи:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu Setting 
      Caption         =   "Настройки"
   End
   Begin VB.Menu Menu_Connect 
      Caption         =   "Соединится"
   End
   Begin VB.Menu Menu_About 
      Caption         =   "О программе..."
   End
End
Attribute VB_Name = "List_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Def_User As String

Private Sub Com_Exit_Click()
 Unload Me
End Sub

Private Sub Com_Ok_Click()
 Dim a As String
 Dim Temp() As String
 Com_Ok.Enabled = False
 List.Enabled = False
 Menu_Connect.Enabled = False
 a = SendCom("Login " & List.List(List.ListIndex))
 If a <> "Error" Then
  Temp() = Split(a, "|")
  a = Temp(0)
  Password.Show vbModal, Me
  If Form_Password = a Then
   ' Запрос сессии
   a = SendCom("SID for " & List.List(List.ListIndex))
   If a <> "Error" Then
    SID = CLng(a)
    ' Запрос типа пользователя
    a = SendCom("Get_Type " & SID)
    If a <> "Error" Then
     User_Type = CLng(a)
     SaveSetting "Zarplata", "Setting", "Def_User", List.List(List.ListIndex)
     C_Main.Show vbModal, Me
    Else
     Add_Status "Ошибка при авторизации"
    End If
   Else
    Add_Status "Ошибка при авторизации"
   End If
  Else
   MsgBox "Пароль неверный!!!", vbCritical, "Ошибка"
   Add_Status "Ошибка при авторизации"
  End If
 Else
  Add_Status "Ошибка при авторизации"
 End If
 Com_Ok.Enabled = True
 List.Enabled = True
 Menu_Connect.Enabled = True
End Sub

Private Sub Form_Load()
 Crypt_Key = Get_Crypt_Key
 Ports = GetSetting("Zarplata", "Setting", "Ports", 1000)
 Host = GetSetting("Zarplata", "Setting", "Host", "127.0.0.1")
 TimeOut = GetSetting("Zarplata", "Setting", "TimeOut", 60)
 Def_User = GetSetting("Zarplata", "Setting", "Def_User", "Admin")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub List_DblClick()
 Call Com_Ok_Click
End Sub

Private Sub Menu_About_Click()
 About.Show vbModal, Me
End Sub

Private Sub Menu_Connect_Click()
 Timer_Connect.Enabled = True
End Sub

Private Sub Setting_Click()
 Sett.Show vbModal, Me
End Sub

Private Sub Timer_Connect_Timer()
 On Error Resume Next
 Dim UsersList As String
 Dim Temp() As String
 Dim i As Long
 Dim D_Num As Long
 If Not Connect(Host, Ports) Then
  Add_Status "Сервер недоступен или все порты заняты, пробуйте позднее."
 Else
  UsersList = SendCom("Get_Users_List")
  If UsersList = "Error" Then
   Add_Status "Ошибка при получении списка пользователей"
  Else
   List.Clear
   Temp() = Split(UsersList, vbCrLf)
   For i = 0 To UBound(Temp()) - 1
    List.AddItem Temp(i)
    If Temp(i) = Def_User Then D_Num = i
   Next
   If Def_User <> vbNullString Then
    List.ListIndex = D_Num
   Else
    List.ListIndex = 0
   End If
   Com_Ok.Enabled = True
   List.Enabled = True
  End If
 End If
 Timer_Connect.Enabled = False
End Sub

Private Sub Winsock_Close()
 Flag_Connect = False
 Add_Status "Отключился"
End Sub

Private Sub Winsock_Connect()
 Add_Status "Подключился"
 Flag_Connect = True
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
 Dim D As String
 Winsock.GetData D
 Add_Status "Получено " & Len(D) & " из " & bytesTotal & " байт..."
 data = data & ZP_DeCode(D)
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Gl_Flag_Err = True
 Add_Status "#" & Number & " - " & Description
 Winsock.Close
End Sub

Private Sub Winsock_SendComplete()
 Add_Status "Отправлено"
End Sub

Private Sub Winsock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
 Add_Status "Отправлено " & Int((bytesRemaining + bytesSent) - bytesSent) & " из " & (bytesRemaining + bytesSent) & " байт..."
End Sub

Function Connect(H As String, Port As Long) As Boolean
 Dim Tmr As Long
 Dim i As Long
 Dim p As String
 Dim Temp() As String
 Dim Flag_Err As Boolean
 Winsock.Close
 Flag_Connect = False
 Gl_Flag_Err = False
 Add_Status "Подключение к " & H & ":" & Port & "..."
 Winsock.Connect H, Ports
 Tmr = Timer()
 Do
  DoEvents
  If Timer() - Tmr > TimeOut Then
   Add_Status "Таймаут при попытке подключится к серверу " & Host & ":" & Port
   Flag_Err = True
  End If
 Loop Until Flag_Connect Or Flag_Err Or Gl_Flag_Err
 If Flag_Connect Then
  p = SendCom("Get_Free_Ports")
  Winsock.Close
  If p <> "Error" Then
   Temp() = Split(p, "|")
   For i = 0 To UBound(Temp())
    Flag_Connect = False
    Gl_Flag_Err = False
    Add_Status "Подключение к " & H & ":" & Port + i & "..."
    Winsock.Connect H, Val(Temp(i))
    Tmr = Timer()
    Do
     DoEvents
     If Timer() - Tmr > TimeOut Then
      Add_Status "Таймаут при попытке подключится к серверу " & Host & ":" & Port + i
      Flag_Err = True
     End If
    Loop Until Flag_Connect Or Flag_Err Or Gl_Flag_Err
    If Flag_Connect Then Exit For
   Next
   Connect = Flag_Connect And Not Flag_Err
  Else
   Connect = False
  End If
 End If
End Function

Function Get_Crypt_Key() ' Ключ шифрования
 On Error GoTo m1
 Dim i As Long
 Dim Temp As String
 Dim k As String
 Open App.Path & "\key.txt" For Input As #1
  Do
   Line Input #1, Temp
   k = k & Temp
  Loop Until EOF(1)
 Close #1
m1:
 If k = vbNullString Then
  Randomize Timer()
  For i = 0 To 1024
   k = k & Chr$(Rnd(1) * 255)
  Next
  Open App.Path & "\key.txt" For Output As #1
   Print #1, k
  Close #1
 End If
 Get_Crypt_Key = k
End Function
