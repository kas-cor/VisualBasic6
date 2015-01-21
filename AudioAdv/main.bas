Attribute VB_Name = "main"
Option Explicit

Global This As Plugin

Public Declare Function FindWindow _
        Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String _
) As Long

Public Declare Function SendMessage _
        Lib "user32" Alias _
        "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByRef lParam As Any _
) As Long

Public Declare Function GetWindowText _
        Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long _
) As Long


Public Declare Sub Sleep _
        Lib "kernel32" ( _
        ByVal dwMilliseconds As Long _
)

Public Sub ld()
    Load frmHidden
End Sub
