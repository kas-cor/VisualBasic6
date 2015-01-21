VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DontWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Календарь не рабочих дней"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "DontWork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   4920
      Top             =   7320
   End
   Begin VB.CommandButton Com_Ok 
      Caption         =   "Применить"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Com_SunDay 
      Caption         =   "Пометить все субботы и воскресенья выходными днями"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   4695
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Закрыть"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   6600
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   11642
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxSelCount     =   366
      MonthColumns    =   4
      MonthRows       =   3
      StartOfWeek     =   21037058
      CurrentDate     =   39612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Caption         =   "Пометьте все выходные дни"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "DontWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_Close_Click()
 Unload Me
End Sub

Private Sub Com_Ok_Click()
 Com_Ok.Enabled = False
 Dim i As Long
 Dim a As String
 Dim Tmp As String
 For i = 1 To 378
  Tmp = MonthView.VisibleDays(i)
  If MonthView.DayBold(CDate(Tmp)) Then
   a = a & Unix_Time(0, 0, CLng(Mid$(Tmp, 1, 2)), CLng(Mid$(Tmp, 4, 2)), CLng(Mid$(Tmp, 7, 4))) & vbCrLf
  End If
 Next
 a = SendCom("Put_DontWorkDays" & vbCrLf & a)
 If a <> "Ok" Then
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Else
  MsgBox "Данные успешно отправлены", vbInformation, "Успешно"
 End If
 Com_Ok.Enabled = True
End Sub

Private Sub Com_SunDay_Click()
 Dim i As Long
 For i = 1 To 378
  If Year(MonthView.VisibleDays(i)) = Year(Date) And Weekday(MonthView.VisibleDays(i), vbMonday) > 5 Then MonthView.DayBold(MonthView.VisibleDays(i)) = True
 Next
End Sub

Private Sub Form_Load()
 MonthView.Value = Date
 Select Case Form_Buh
  Case True
   Label1.Caption = "Пометьте все не рабочие дни"
  Case False
   Label1.Caption = "Выходные дни"
   Com_SunDay.Visible = False
   Com_Ok.Visible = False
 End Select
End Sub

Private Sub MonthView_DateClick(ByVal DateClicked As Date)
 If Form_Buh Then MonthView.DayBold(DateClicked) = Not MonthView.DayBold(DateClicked)
End Sub

Private Sub Timer_Timer()
 Dim a As String
 Dim i As Long
 Dim Temp() As String
 a = SendCom("Get_DontWorkDays")
 If a <> "Error" Then
  Temp() = Split(a, vbCrLf)
  For i = 0 To UBound(Temp()) - 1
   MonthView.DayBold(CDate(Mid$(Time_Unix(CLng(Temp(i))), 7))) = True
  Next
 Else
  MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 End If
 If Form_Buh Then MonthView.Enabled = True
 Com_SunDay.Enabled = True
 Com_Ok.Enabled = True
 Com_Close.Enabled = True
 Timer.Enabled = False
End Sub
