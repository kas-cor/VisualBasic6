VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2370
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   54984706
      CurrentDate     =   39615
   End
   Begin VB.CommandButton Com_Print 
      Caption         =   "��� ������"
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton Com_Report 
      Caption         =   "�������"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Com_Close 
      Cancel          =   -1  'True
      Caption         =   "�������"
      Height          =   375
      Left            =   3540
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin ComctlLib.ListView ListView 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "date"
         Object.Tag             =   ""
         Text            =   "����"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "TimeStart"
         Object.Tag             =   ""
         Text            =   "������"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Obed"
         Object.Tag             =   ""
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "TimeEnd"
         Object.Tag             =   ""
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "TimeAll"
         Object.Tag             =   ""
         Text            =   "����������"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   "Zar"
         Object.Tag             =   ""
         Text            =   "����������"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComCtl2.DTPicker DT_Date 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   54984705
      CurrentDate     =   39603
   End
   Begin MSComCtl2.DTPicker DT_Date 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   54984705
      CurrentDate     =   39603
   End
   Begin VB.Label Label6 
      Caption         =   "������ � ���:"
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
      TabIndex        =   15
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   6840
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "����� ����������:"
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
      TabIndex        =   10
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "����� ����������:"
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
      TabIndex        =   9
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label23 
      Caption         =   "������:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "�"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label25 
      Caption         =   "��"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DWD_Max As Long
Dim DWD_Date(0 To 1000) As Long
Dim Work_Hr As Long ' ������� �����

Private Sub Com_Close_Click()
 Unload Me
End Sub

Private Sub Com_Print_Click()
 File_Run App.Path & "\report.html"
End Sub

Private Sub Com_Report_Click()
 Com_Report.Enabled = False
 Com_Print.Enabled = False
 Com_Close.Enabled = False
 Dim a As String
 Dim s As Long
 Dim e As Long
 Dim i As Long
 Dim Tmp As String
 Dim Temp1() As String
 Dim Temp2() As String
 Dim R_Max As Long '           ����� �������
 Dim RS(0 To 5000) As Long '   ������
 Dim RO_S(0 To 5000) As Long ' ������ �����
 Dim RO_E(0 To 5000) As Long ' ����� �����
 Dim Obed_Txt As String '      ���� �����
 Dim RE(0 To 5000) As Long '   ����
 Dim Ras_Min As Long '         ����� �� ����
 Dim Ras_Txt As String '       ����� � �����
 Dim Ras_Min_Sum As Long '     ����� �����
 Dim Stav(1 To 12) As Double ' ������ � ���
 Dim ZP As Double '            ��������
 Dim Shablon As String '       ������
 Dim HTML_Shab As String '     ������������ ������
 Dim TempShab() As String
 a = SendCom("Data_For_User " & Form_User)
 If a = "Error" Then
  MsgBox "������ ��������! ���������� ��� ���.", vbCritical, "������"
  Exit Sub
 Else
  Temp1() = Split(a, "|")
  If Temp1(0) = "1" Then ' ������������ ������
   For i = 1 To 12
    Stav(i) = CDbl(Temp1(1)) / Get_WorkDays(i) / Work_Hr
   Next
  ElseIf Temp1(0) = "0" Then ' ����������� ������
   For i = 1 To 12
    Stav(i) = CDbl(Temp1(2))
   Next
  End If
  Erase Temp1()
 End If
 s = Unix_Time(0, 0, DT_Date(0).Day, DT_Date(0).Month, DT_Date(0).Year)
 e = Unix_Time(0, 0, DT_Date(1).Day, DT_Date(1).Month, DT_Date(1).Year)
 If s > e Then MsgBox "��������� ���� ������ ��������", vbCritical, "������": Exit Sub
 a = SendCom("Report|" & Form_User & "|" & s & "|" & e)
 If a = "Error" Then
  MsgBox "������ ��������! ���������� ��� ���.", vbCritical, "������"
 ElseIf a = "Ok" Then
  MsgBox "�� ������ ������ ������� ������ ���.", vbInformation, "�����"
 Else
  Temp1() = Split(a, vbCrLf)
  For i = 0 To UBound(Temp1()) - 1
   Temp2() = Split(Temp1(i), ",")
   RS(i) = CLng(Temp2(0))
   RO_S(i) = CLng(Temp2(1))
   RO_E(i) = CLng(Temp2(2))
   RE(i) = CLng(Temp2(3))
  Next
  R_Max = i
  Shablon = SendCom("Get_Shablon1")
  If Shablon = "Error" Then
   MsgBox "����������� ���� �������! ������������ ������ �� ���������!", vbCritical, "������"
  Else
   TempShab() = Split(Shablon, "<!--{split}-->")
   If UBound(TempShab()) <> 2 Then
    MsgBox "���� ������� �� ������� �������! ������������ ������ �� ���������!", vbCritical, "������"
    Shablon = "Error"
   End If
  End If
  Open App.Path & "\report.html" For Output As #1
   ' �����
   If Shablon = "Error" Then
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<meta http-equiv='Content-Type' content='text/html; charset=windows-1251'>"
    Print #1, "<meta http-equiv='Content - Language' content='ru'>"
    Print #1, "<title>����� �� ���������� �����</title>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<h2><center><b>����� �� ���������� �����</b></center></h2>"
    Print #1, "<p align='Left'><b>��������:</b>&nbsp;" & Form_User & "<br>"
    Print #1, "<b>������:</b>&nbsp;�&nbsp;" & Mid$(Time_Unix(s), 7) & "&nbsp;��&nbsp;" & Mid$(Time_Unix(e), 7) & "</p>"
    Print #1, "<table border='1' cellspacing='1' style='border-collapse: collapse' width='100%'>"
    Print #1, "<tr>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>����</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>���� ������</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>������</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>����</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>����</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>����������</b></td>"
    Print #1, "<td align='center' bgcolor='#C0C0C0' width='14%'><b>����������</b></td>"
    Print #1, "</tr>"
   Else
    HTML_Shab = Replace(TempShab(0), "{#user_name#}", Form_User, , , vbTextCompare)
    HTML_Shab = Replace(HTML_Shab, "{#date_start#}", Mid$(Time_Unix(s), 7), , , vbTextCompare)
    HTML_Shab = Replace(HTML_Shab, "{#date_end#}", Mid$(Time_Unix(e), 7), , , vbTextCompare)
    Print #1, HTML_Shab
   End If
   ListView.ListItems.Clear
   ' ��������� �����
   For i = 0 To R_Max - 1
    Tmp = Time_Unix(RS(i))
    Ras_Min = (RE(i) - RS(i)) - (RO_E(i) - RO_S(i))
    Ras_Min_Sum = Ras_Min_Sum + Ras_Min
    ZP = ZP + Ras_Min * (Stav(CLng(Mid$(Tmp, 10, 2))) / 60)
    Ras_Txt = Unix_Text(Ras_Min) ' ���������� �����
    Obed_Txt = Unix_Text(RO_E(i) - RO_S(i)) ' ���� �����
    ListView.ListItems.Add , "id" & i, Mid$(Tmp, 7)
    ListView.ListItems("id" & i).SubItems(1) = Mid$(Time_Unix(RS(i)), 1, 5) ' ������
    ListView.ListItems("id" & i).SubItems(2) = Obed_Txt '                     ����
    ListView.ListItems("id" & i).SubItems(3) = Mid$(Time_Unix(RE(i)), 1, 5) ' ����
    ListView.ListItems("id" & i).SubItems(4) = Ras_Txt '                      ����������
    ListView.ListItems("id" & i).SubItems(5) = Format$(Ras_Min * (Stav(CLng(Mid$(Tmp, 10, 2))) / 60), "0.00") ' ����������
    If Shablon = "Error" Then
     Print #1, "<tr>"
     Print #1, "<td align='Center' width='14%'>" & Mid$(Time_Unix(RS(i)), 7) & "</td>"
     Print #1, "<td align='Center' width='14%'>" & WeekdayName(Weekday(CDate(Mid$(Time_Unix(RS(i)), 7)))) & "</td>"
     Print #1, "<td align='Center' width='14%'>" & Mid$(Time_Unix(RS(i)), 1, 5) & "</td>"
     Print #1, "<td align='Center' width='14%'>" & Obed_Txt & "</td>"
     Print #1, "<td align='Center' width='14%'>" & Mid$(Time_Unix(RE(i)), 1, 5) & "</td>"
     Print #1, "<td align='Center' width='14%'>" & Ras_Txt & "</td>"
     Print #1, "<td align='Center' width='14%'>" & Format$(Ras_Min * (Stav(CLng(Mid$(Tmp, 10, 2))) / 60), "0.00") & "</td>"
     Print #1, "</tr>"
    Else
     HTML_Shab = Replace(TempShab(1), "{#date#}", Mid$(Time_Unix(RS(i)), 7), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#week_day_name#}", WeekdayName(Weekday(CDate(Mid$(Time_Unix(RS(i)), 7)))), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#time_start#}", Mid$(Time_Unix(RS(i)), 1, 5), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#dinner#}", Obed_Txt, , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#time_end#}", Mid$(Time_Unix(RE(i)), 1, 5), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#time#}", Ras_Txt, , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#pay#}", Format$(Ras_Min * (Stav(CLng(Mid$(Tmp, 10, 2))) / 60), "0.00"), , , vbTextCompare)
     Print #1, HTML_Shab
    End If
   Next
   ' ������
   Ras_Txt = Unix_Text(Ras_Min_Sum)
   Label3.Caption = Ras_Txt
   Label5.Caption = Format$(ZP / (Ras_Min_Sum / 60), "0.00")
   Label4.Caption = Format$(ZP, "0.00")
   If Shablon = "Error" Then
    Print #1, "<tr>"
    Print #1, "<td align='Right' width='86%' colspan='6'><b>����� ����������:</b></td>"
    Print #1, "<td width='16%' align='Center'>" & Ras_Txt & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr>"
    Print #1, "<td align='Right' width='86%' colspan='6'><b>������ � ���:</b></td>"
    Print #1, "<td width='16%' align='Center'>" & Format$(ZP / (Ras_Min_Sum / 60), "0.00") & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr>"
    Print #1, "<td align='Right' width='86%' colspan='6'><b>����� ����������:</b></td>"
    Print #1, "<td width='16%' align='Center'>" & Format$(ZP, "0.00") & "</td>"
    Print #1, "</tr>"
    Print #1, "</table>"
    Print #1, "</body>"
    Print #1, "</html>"
   Else
    HTML_Shab = Replace(TempShab(2), "{#time_all#}", Ras_Txt, , , vbTextCompare)
    HTML_Shab = Replace(HTML_Shab, "{#stavka#}", Format$(ZP / (Ras_Min_Sum / 60), "0.00"), , , vbTextCompare)
    HTML_Shab = Replace(HTML_Shab, "{#pay_all#}", Format$(ZP, "0.00"), , , vbTextCompare)
    Print #1, HTML_Shab
   End If
  Close #1
 End If
 Com_Report.Enabled = True
 Com_Print.Enabled = True
 Com_Close.Enabled = True
End Sub

Private Sub Form_Load()
 Dim a As String
 Dim i As Long
 Dim Temp() As String
 DT_Date(0).Day = Day(Date)
 DT_Date(0).Month = Month(Date)
 DT_Date(0).Year = Year(Date)
 DT_Date(1).Day = Day(Date)
 DT_Date(1).Month = Month(Date)
 DT_Date(1).Year = Year(Date)
 ' �������� �� ������� ���
 a = SendCom("Get_DontWorkDays")
 If a <> "Error" Then
  Temp() = Split(a, vbCrLf)
  DWD_Max = UBound(Temp()) - 1
  For i = 0 To DWD_Max
   DWD_Date(i) = CLng(Temp(i))
  Next
 Else
  MsgBox "������ ��������! ���������� ��� ���.", vbCritical, "������"
 End If
 Erase Temp()
 a = SendCom("Get_Buh_Sett")
 If a = "Error" Then
  MsgBox "������ ��������! ���������� ��� ���.", vbCritical, "������"
 Else
  Temp() = Split(a, "|")
  Work_Hr = CLng(Temp(4))
 End If
End Sub

Private Function Get_WorkDays(m As Long) As Long
 Dim i As Long
 Dim t As Long
 Dim w As Long
 Dim Flag As Boolean
 Dim Tmp As String
 MonthView.Value = "01." & IIf(m < 10, "0", "") & m & "." & Year(Date)
 For t = 1 To 42
  Tmp = MonthView.VisibleDays(t)
  If CLng(Mid$(Tmp, 4, 2)) = m Then
   Flag = False
   For i = 0 To DWD_Max
    If Unix_Time(0, 0, CLng(Mid$(Tmp, 1, 2)), m, Year(Date)) = DWD_Date(i) Then Flag = True: Exit For
   Next
   If Not Flag Then w = w + 1
  End If
 Next
 Get_WorkDays = w
End Function
