VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Report_All 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отчет"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Report_All.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Com_Print 
      Caption         =   "Для печати"
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Com_Report 
      Caption         =   "Создать"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Com_Close 
      Cancel          =   -1  'True
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   3540
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2370
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   21757954
      CurrentDate     =   39615
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "User"
         Object.Tag             =   ""
         Text            =   "Работник"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "TimeAll"
         Object.Tag             =   ""
         Text            =   "Отработал"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Stav"
         Object.Tag             =   ""
         Text            =   "Ставка в час"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Zar"
         Object.Tag             =   ""
         Text            =   "Заработал"
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
      Format          =   21757953
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
      Format          =   21757953
      CurrentDate     =   39603
   End
   Begin VB.Label Label23 
      Caption         =   "Период:"
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
      Caption         =   "с"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label25 
      Caption         =   "по"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Report_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DWD_Max As Long
Dim DWD_Date(0 To 1000) As Long
Dim Work_Hr As Long ' Рабочих часов

Dim Usr_Max As Long '          Всего работников
Dim Usr(0 To 1000) As String ' Работник

Private Sub Com_Close_Click()
 Unload Me
End Sub

Private Sub Com_Print_Click()
 File_Run App.Path & "\report_all.html"
End Sub

Private Sub Com_Report_Click()
 Com_Report.Enabled = False
 Com_Print.Enabled = False
 Com_Close.Enabled = False
 Dim a As String
 Dim s As Long
 Dim e As Long
 Dim i As Long
 Dim u As Long
 Dim Tmp As String
 Dim Temp1() As String
 Dim Temp2() As String
 Dim Ras_Min As Long '         Минут за день
 Dim Ras_Min_Sum As Long '     Минут всего
 Dim Stav(1 To 12) As Double ' Ставка в час
 Dim ZP As Double '            Зарплата
 Dim Shablon As String '       Шаблон
 Dim HTML_Shab As String '     Обработанный шаблон
 Dim TempShab() As String
 s = Unix_Time(0, 0, DT_Date(0).Day, DT_Date(0).Month, DT_Date(0).Year)
 e = Unix_Time(0, 0, DT_Date(1).Day, DT_Date(1).Month, DT_Date(1).Year)
 If s > e Then MsgBox "Начальная дата больше конечной", vbCritical, "Ошибка": Exit Sub
 Shablon = SendCom("Get_Shablon2")
 If Shablon = "Error" Then
  MsgBox "Отсутствует файл шаблона! Используется шаблон по умолчанию!", vbCritical, "Ошибка"
 Else
  TempShab() = Split(Shablon, "<!--{split}-->")
  If UBound(TempShab()) <> 2 Then
   MsgBox "Файл шаблона не верного формата! Используется шаблон по умолчанию!", vbCritical, "Ошибка"
   Shablon = "Error"
  End If
 End If
 Open App.Path & "\report_all.html" For Output As #1
  ' Шапка
  ListView.ListItems.Clear
  If Shablon = "Error" Then
   Print #1, "<html>"
   Print #1, "<head>"
   Print #1, "<meta http-equiv='Content-Type' content='text/html; charset=windows-1251'>"
   Print #1, "<meta http-equiv='Content - Language' content='ru'>"
   Print #1, "<title>Отчет по заработной плате по всем работникам</title>"
   Print #1, "</head>"
   Print #1, "<body>"
   Print #1, "<h2><center><b>Отчет по заработной плате по всем работникам</b></center></h2>"
   Print #1, "<b>Период:</b>&nbsp;с&nbsp;" & Mid$(Time_Unix(s), 7) & "&nbsp;по&nbsp;" & Mid$(Time_Unix(e), 7) & "</p>"
   Print #1, "<table border='1' cellspacing='1' style='border-collapse: collapse' width='100%'>"
   Print #1, "<tr>"
   Print #1, "<td align='center' bgcolor='#C0C0C0' width='25%'><b>Работник</b></td>"
   Print #1, "<td align='center' bgcolor='#C0C0C0' width='25%'><b>Отработал</b></td>"
   Print #1, "<td align='center' bgcolor='#C0C0C0' width='25%'><b>Ставка в час</b></td>"
   Print #1, "<td align='center' bgcolor='#C0C0C0' width='25%'><b>Заработал</b></td>"
   Print #1, "</tr>"
  Else
   HTML_Shab = Replace(TempShab(0), "{#date_start#}", Mid$(Time_Unix(s), 7), , , vbTextCompare)
   HTML_Shab = Replace(HTML_Shab, "{#date_end#}", Mid$(Time_Unix(e), 7), , , vbTextCompare)
   Print #1, HTML_Shab
  End If
  For u = 0 To Usr_Max
   a = SendCom("Data_For_User " & Usr(u))
   If a = "Error" Then
    MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
    Exit Sub
   Else
    Temp1() = Split(a, "|")
    If Temp1(0) = "1" Then ' Динамическая ставка
     For i = 1 To 12
      Stav(i) = CDbl(Temp1(1)) / Get_WorkDays(i) / Work_Hr
     Next
    ElseIf Temp1(0) = "0" Then ' Статическая ставка
     For i = 1 To 12
      Stav(i) = CDbl(Temp1(2))
     Next
    End If
    Erase Temp1()
   End If
   a = SendCom("Report|" & Usr(u) & "|" & s & "|" & e)
   If a = "Error" Then
    MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
   ElseIf a = "Ok" Then
    MsgBox "За данный период времени данных для работника " & Usr(u) & " нет.", vbInformation, "Отчет"
   Else
    Ras_Min_Sum = 0
    ZP = 0
    Temp1() = Split(a, vbCrLf)
    For i = 0 To UBound(Temp1()) - 1
     Temp2() = Split(Temp1(i), ",")
     Tmp = Time_Unix(CLng(Temp2(3)))
     Ras_Min = (CLng(Temp2(3)) - CLng(Temp2(0))) - (CLng(Temp2(2)) - CLng(Temp2(1)))
     Ras_Min_Sum = Ras_Min_Sum + Ras_Min
     ZP = ZP + Ras_Min * (Stav(CLng(Mid$(Tmp, 10, 2))) / 60)
     Erase Temp2()
    Next
    Erase Temp1()
    ' Табличная часть
    ListView.ListItems.Add , "id" & u, Usr(u) ' Работник
    ListView.ListItems("id" & u).SubItems(1) = Unix_Text(Ras_Min_Sum)  '              Отработано
    ListView.ListItems("id" & u).SubItems(2) = Format$(ZP / (Ras_Min_Sum / 60), "0.00") ' Ставка в час
    ListView.ListItems("id" & u).SubItems(3) = Format$(ZP, "0.00") '                      Заработано
    If Shablon = "Error" Then
     Print #1, "<tr>"
     Print #1, "<td align='Center' width='25%'>" & Usr(u) & "</td>"
     Print #1, "<td align='Center' width='25%'>" & Unix_Text(Ras_Min_Sum) & "</td>"
     Print #1, "<td align='Center' width='25%'>" & Format$(ZP / (Ras_Min_Sum / 60), "0.00") & "</td>"
     Print #1, "<td align='Center' width='25%'>" & Format$(ZP, "0.00") & "</td>"
     Print #1, "</tr>"
    Else
     HTML_Shab = Replace(TempShab(1), "{#user_name#}", Usr(u), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#time_all#}", Unix_Text(Ras_Min_Sum), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#stavka#}", Format$(ZP / (Ras_Min_Sum / 60), "0.00"), , , vbTextCompare)
     HTML_Shab = Replace(HTML_Shab, "{#pay_all#}", Format$(ZP, "0.00"), , , vbTextCompare)
     Print #1, HTML_Shab
    End If
   End If
  Next
  ' Подвал
  If Shablon = "Error" Then
   Print #1, "</table>"
   Print #1, "</body>"
   Print #1, "</html>"
  Else
   Print #1, TempShab(2)
  End If
 Close #1
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
 ' Получает не рабочие дни
 a = SendCom("Get_DontWorkDays")
 If a <> "Error" Then
  Temp() = Split(a, vbCrLf)
  DWD_Max = UBound(Temp()) - 1
  For i = 0 To DWD_Max
   DWD_Date(i) = CLng(Temp(i))
  Next
 Else
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 End If
 Erase Temp()
 a = SendCom("Get_Buh_Sett")
 If a = "Error" Then
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Else
  Temp() = Split(a, "|")
  Work_Hr = CLng(Temp(4))
 End If
 a = SendCom("Get_Users_List")
 If a = "Error" Then
   MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Else
  Temp() = Split(a, vbCrLf)
  Usr_Max = UBound(Temp()) - 3
  For i = 0 To Usr_Max
   Usr(i) = Temp(i)
  Next
 End If
 Erase Temp()
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

