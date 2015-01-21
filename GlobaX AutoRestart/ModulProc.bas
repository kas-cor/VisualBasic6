Attribute VB_Name = "ModulProc"
Option Explicit

Const TH32CS_SNAPPROCESS As Long = 2&
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)

Public Function GX_Present() As Boolean
 Dim hSnapShot As Long
 Dim uProcess As PROCESSENTRY32
 Dim r As Long
 hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
 If hSnapShot = 0 Then
  Exit Function
 End If
 uProcess.dwSize = Len(uProcess)
 r = ProcessFirst(hSnapShot, uProcess)
 Do While r
  If Mid(Trim(LCase(uProcess.szExeFile)), 1, 17) = "globax_daemon.exe" Then
   GX_Present = True
   Exit Do
  End If
  r = ProcessNext(hSnapShot, uProcess)
 Loop
 Call CloseHandle(hSnapShot)
End Function
