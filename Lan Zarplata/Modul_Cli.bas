Attribute VB_Name = "Modul"
Option Explicit

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As SINITCOMMONCONTROLSEX) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Type SINITCOMMONCONTROLSEX
   dwSize As Long
   dwICC As Long
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Crypt_Key As String * 1024 '        ���� ��� ����������

Public Flag_Connect As Boolean ' ���� ��������
Public Gl_Flag_Err As Boolean '  ���� ������
Public data As String '          ������ �� �������

Public Host As String '          ����� �������
Public Ports As Long '           ����� �������
Public TimeOut As Long '         ����� �������

Public Form_Password As String ' ������ �� ����� �����
Public Form_User As String '     ������������ �� �����
Public Form_Buh As Boolean '     � ��������� ������� ���������?

Public SID As Long '             �� ������ ������������
Public User_Type As Long '       ��� ������������

Sub Main()
 Call InitXPStyle
 List_Users.Show
End Sub

Public Sub InitXPStyle()
 Dim InitCtrls As SINITCOMMONCONTROLSEX
 On Error Resume Next
 With InitCtrls
  .dwSize = LenB(InitCtrls)
  .dwICC = ICC_USEREX_CLASSES
 End With
 InitCommonControlsEx InitCtrls
End Sub

Public Function SendCom(Com As String) As String
 If List_Users.Winsock.State = sckClosed Then
  SendCom = "Error"
  Exit Function
 End If
 Dim Tmr As Long
 Dim Flag_Err As Boolean
 data = vbNullString
 List_Users.Winsock.SendData ZP_EnCode(Com & "." & vbCrLf)
 Tmr = Timer()
 Do
  DoEvents
  If Timer() - Tmr > TimeOut Then
   Add_Status "������� ��� ������� �������� ������"
   Flag_Err = True
  End If
 Loop Until InStr(1, data, "." & vbCrLf) <> 0 Or Flag_Err
 If Not Flag_Err Then
  SendCom = Mid$(data, 1, Len(data) - 3)
 Else
  SendCom = "Error"
 End If
End Function

Public Function RC4(data As String) As String
 Dim Key(0 To 255) As Long
 Dim box(0 To 255) As Long
 Dim cipher As String
 Dim pwd_length As Long
 Dim data_length As Long
 Dim J As Long
 Dim a As Long
 Dim i As Long
 Dim k As Long
 Dim Tmp As Long
 Dim txt As String
 Dim pwd As String
 pwd = Crypt_Key
 pwd_length = Len(pwd)
 data_length = Len(data)
 For i = 0 To 255
  Key(i) = Asc(Mid$(pwd, (i Mod pwd_length) + 1, 1))
  box(i) = i
 Next
 J = 0
 For i = 0 To 255
  J = (J + box(i) + Key(i)) Mod 256
  Tmp = box(i)
  box(i) = box(J)
  box(J) = Tmp
 Next
 a = 0
 J = 0
 For i = 1 To data_length
  a = (a + 1) Mod 256
  J = (J + box(a)) Mod 256
  Tmp = box(a)
  box(a) = box(J)
  box(J) = Tmp
  k = box((box(a) + box(J)) Mod 256)
  cipher = cipher & Chr$(Asc(Mid$(data, i, 1)) Xor k)
 Next
 RC4 = cipher
End Function

Function ZP_DeCode(txt As String) As String
 Dim s As String
 Dim c As String
 Dim rs1 As String
 Dim rs2 As String
 Dim r1 As Long
 Dim r2 As Long
 Dim i As Long
 For i = 1 To Len(txt) Step 2
  rs1 = Mid$(txt, i, 1)
  rs2 = Mid$(txt, i + 1, 1)
  If Asc(rs1) - 48 > -1 And Asc(rs1) - 48 < 10 Then r1 = Asc(rs1) - 48
  If Asc(rs1) - 55 > 9 And Asc(rs1) - 55 < 16 Then r1 = Asc(rs1) - 55
  If Asc(rs2) - 48 > -1 And Asc(rs2) - 48 < 10 Then r2 = Asc(rs2) - 48
  If Asc(rs2) - 55 > 9 And Asc(rs2) - 55 < 16 Then r2 = Asc(rs2) - 55
  s = s & Chr$(r1 * 16 + r2)
 Next
 ZP_DeCode = s
End Function

Function ZP_EnCode(txt As String) As String
 Dim s As String
 Dim c As String
 Dim i As Long
 For i = 1 To Len(txt)
  c = Hex$(Asc(Mid$(txt, i, 1)))
  If Len(c) = 1 Then c = "0" & c
  s = s & c
 Next
 ZP_EnCode = s
End Function

Public Function Unix_Time(Hr As Long, Min As Long, Day As Long, Mount As Long, Year As Long) As Long
 Dim u As Long
 Dim y As Long
 y = Year
 If y < 1970 Then y = 0
 y = y - 1970
 u = y * 599040 ' ����� � ����
 u = u + Mount * 46080 ' ����� � ������
 u = u + Day * 1440 ' ����� � ������
 u = u + Hr * 60 ' ����� � ����
 u = u + Min
 Unix_Time = u
End Function

Public Function Time_Unix(u As Long) As String
 Dim Tmp As Long
 Dim Hr As Long
 Dim Hr_Zero As String
 Dim Min As Long
 Dim Min_Zero As String
 Dim Day As Long
 Dim Day_Zero As String
 Dim Mount As Long
 Dim Mount_Zero As String
 Dim Year As Long
 Tmp = u
 Year = Int(Tmp / 599040) ' ����� � ����
 Tmp = Tmp - Year * 599040
 Mount = Int(Tmp / 46080) ' ����� � ������
 Tmp = Tmp - Mount * 46080
 Day = Int(Tmp / 1440) ' ����� � ���
 Tmp = Tmp - Day * 1440
 Hr = Int(Tmp / 60) ' ����� � ����
 Tmp = Tmp - Hr * 60
 Min = Tmp
 If Hr < 10 Then Hr_Zero = "0"
 If Min < 10 Then Min_Zero = "0"
 If Day < 10 Then Day_Zero = "0"
 If Mount < 10 Then Mount_Zero = "0"
 Time_Unix = Hr_Zero & Hr & ":" & Min_Zero & Min & " " & Day_Zero & Day & "." & Mount_Zero & Mount & "." & Year + 1970
End Function

Public Sub Add_Status(txt As String)
 List_Users.StatusBar.Panels(1).Text = txt
End Sub

Sub File_Run(File As String)
 ShellExecute C_Main.hwnd, vbNullString, File, vbNullString, Mid$(App.Path, 1, 3), SW_SHOWNORMAL
End Sub

Public Function Unix_Text(u As Long) As String
 ' ������� UNIX ������� � �����
 Dim Tmp As String
 Tmp = CStr(Int(u / 60)) & " ���. "
 If Tmp = "0 ���. " Then Tmp = vbNullString
 Tmp = Tmp & CStr(u - Int(u / 60) * 60) & " ���."
 If Right$(Tmp, 7) = " 0 ���." Then Tmp = Mid$(Tmp, 1, Len(Tmp) - 7)
 Unix_Text = Tmp
End Function
