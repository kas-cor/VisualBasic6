VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHidden 
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wnsServer 
      Index           =   0
      Left            =   1425
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmHidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'sysstem messages
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111

'Winamp Messages
Private Const WM_Raise_Volume = 40058       'increase 1%
Private Const WM_Lower_Volume = 40059       'decrease 1%
Private Const WM_Close_Winamp = 40001
Private Const WM_Previous = 40044
Private Const WM_Next = 40048
Private Const WM_Play = 40045
Private Const WM_Pause_Unpause = 40046
Private Const WM_Stop = 40047
Private Const WM_Toggle_Shuffle = 40023
Private Const WA_SETVOLUME = 122

Dim Response As String
Dim Connections As Long

Private Sub Form_Load()
On Error Resume Next
  Me.Visible = False
  wnsServer(0).Protocol = sckTCPProtocol
  wnsServer(0).LocalPort = 806
  wnsServer(0).Listen
End Sub

Private Sub wnsServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  If Index = 0 Then
     Connections = Connections + 1
     Load wnsServer(Connections) 'Загрузка нового сокета
     wnsServer(Connections).LocalPort = 0  'используем динамический порт
     wnsServer(Connections).Accept requestID 'принимаем коннект
  End If
  DoEvents
End Sub

Private Sub wnsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
        Dim hwnd As Long
        
        hwnd = FindWindow("Winamp v1.x", vbNullString)
        'если к нам подконектились и если у нас присутствует винамп,  ждем команду для отправки
        If bytesTotal <> 0 Then
            
            wnsServer(Index).GetData Response 'получаем данные
                    
            'если нет винампа то можно только выходить
            If hwnd = 0 Then
               Exit Sub
            End If
            
            'Next Track
            If InStr(1, Response, "next", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Next, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Previous Track
            If InStr(1, Response, "previous", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Previous, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Play
            If InStr(1, Response, "play", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Play, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Stop
            If InStr(1, Response, "stop", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Stop, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Shuffle
            If InStr(1, Response, "shuffle", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Toggle_Shuffle, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Pause/UnPause
            If InStr(1, Response, "pause", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Pause_Unpause, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Close
            If InStr(1, Response, "close", vbTextCompare) <> 0 Then
               SendMessage hwnd, WM_COMMAND, WM_Close_Winamp, vbNull
               CloseSocket Index
               Exit Sub
            End If
            'Volume inc
            If InStr(1, Response, "+", vbTextCompare) <> 0 Then
               If Response = "+" Then Response = "+1"
               Volume hwnd, CInt(Mid$(Response, InStr(1, Response, "+") + 1, 3)), 1
               CloseSocket Index
               Exit Sub
            End If
            'Volume dec
            If InStr(1, Response, "-", vbTextCompare) <> 0 Or InStr(1, Response, "0", vbTextCompare) <> 0 Then
               If Mid$(Response, InStr(1, Response, "-") + 1, 3) < "A" Then
                  If Response = "-" Then Response = "-1"
                  Volume hwnd, CInt(Mid$(Response, InStr(1, Response, "-") + 1, 3)), -1
                  CloseSocket Index
                  Exit Sub
               End If
            End If
      End If 'bytes
End Sub

Private Sub Volume(hwnd As Long, percent As Integer, incdec As Long)
    Dim i As Long
        For i = 0 To percent - 1
            Select Case incdec
                    Case -1
                        SendMessage hwnd, WM_COMMAND, WM_Lower_Volume, vbNull
                    Case 1
                        SendMessage hwnd, WM_COMMAND, WM_Raise_Volume, vbNull
            End Select
        Next i
End Sub

Private Sub wnsServer_SendComplete(Index As Integer)
    Sleep 1000
    CloseSocket Index
End Sub

Private Sub CloseSocket(Index As Integer)
    wnsServer(Index).Close
    Unload wnsServer(Connections)
    Connections = Connections - 1
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
   If Connections > 0 Then
        For i = Connections To 1
            wnsServer(i).Close
            Unload wnsServer(Connections)
        Next i
  End If
  DoEvents
End Sub
