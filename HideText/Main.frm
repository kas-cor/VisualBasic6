VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Скрытие текста в программе"
   ClientHeight    =   4695
   ClientLeft      =   7305
   ClientTop       =   6360
   ClientWidth     =   6375
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Копировать в буфер"
      Height          =   375
      Left            =   3240
      TabIndex        =   40
      Top             =   4200
      Width           =   2415
   End
   Begin VB.ComboBox Combo_Proc 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Com_Gen 
      Caption         =   "Сгенирировать"
      Height          =   375
      Left            =   720
      TabIndex        =   37
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text_Result 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   35
      Top             =   3360
      Width           =   6135
   End
   Begin VB.TextBox Text_Source 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   34
      Top             =   2280
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Символы"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   29
         Left            =   5520
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   28
         Left            =   4920
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   27
         Left            =   4320
         TabIndex        =   30
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   26
         Left            =   3720
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   25
         Left            =   3120
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   24
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   23
         Left            =   1920
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   22
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   21
         Left            =   720
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   20
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   19
         Left            =   5520
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   18
         Left            =   4920
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   17
         Left            =   4320
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   16
         Left            =   3720
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   15
         Left            =   3120
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   14
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   13
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   12
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   11
         Left            =   720
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   9
         Left            =   5520
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   8
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   7
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   6
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   5
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   4
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "!"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      ItemData        =   "Main.frx":000C
      Left            =   2880
      List            =   "Main.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "На сколько процентов заполнить текст доп. символами:"
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
      TabIndex        =   38
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "Релутьтат:"
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
      TabIndex        =   36
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Исходный текст:"
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
      TabIndex        =   33
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Дополнительные символы:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rnd_S_Max As Integer
Dim Rnd_S(0 To 29) As String

Dim Flag_ReRnd As Boolean

Private Sub Com_Gen_Click()
 If Text_Source.Text = "" Then MsgBox "Введите исходный текст", vbInformation, "Ошибка!": Exit Sub
 Dim Txt_Sou As String
 Dim Txt_Res As String
 Dim Tmp_Txt As String
 Dim Num_Rnd As String
 Dim Txt_Rep As String
 Dim i As Integer
 Dim t As Integer
 Txt_Sou = Text_Source.Text
 Flag_ReRnd = False
 For i = 1 To Len(Txt_Sou)
  For t = 0 To Rnd_S_Max
   If Mid(Txt_Sou, i, 1) = Rnd_S(t) Then Flag_ReRnd = True
  Next
 Next
 If Flag_ReRnd Then Call Combo_Click: Exit Sub
 Randomize Timer
 For i = 1 To Len(Txt_Sou)
  If Int(Rnd(1) * 100) < (Combo_Proc.ListIndex + 1) * 10 Then
   Txt_Res = Txt_Res & Rnd_S(Int(Rnd(1) * Rnd_S_Max))
  End If
  Txt_Res = Txt_Res & Mid(Txt_Sou, i, 1)
 Next
 For i = 1 To Len(Txt_Res)
  Tmp_Txt = Tmp_Txt & "Chr(" & Asc(Mid(Txt_Res, i, 1)) & ")&"
 Next
 Txt_Res = Mid(Tmp_Txt, 1, Len(Tmp_Txt) - 1)
 Num_Rnd = "" & Int(Rnd(1) * 1000 + 1)
 If Combo.ListIndex <> 0 Then
  Txt_Rep = "Replace(Txt" & Num_Rnd & "," & Chr(34) & Rnd_S(0) & Chr(34) & "," & Chr(34) & Chr(34) & ")"
  For i = 1 To Rnd_S_Max
   If Rnd_S(i) = Chr(34) Then Rnd_S(i) = Chr(34) & Chr(34)
   Txt_Rep = "Replace(" & Txt_Rep & "," & Chr(34) & Rnd_S(i) & Chr(34) & "," & Chr(34) & Chr(34) & ")"
  Next
  Txt_Rep = "<переменная>=" & Txt_Rep
 End If
 Text_Result.Text = "' BEGIN hidden text '" & Text_Source.Text & "'" & vbCrLf & _
 "Dim Txt" & Num_Rnd & " As String" & vbCrLf & _
 "Txt" & Num_Rnd & "=" & Txt_Res & vbCrLf & Txt_Rep & vbCrLf & _
 "' END hidden text '" & Text_Source.Text & "'"
End Sub

Private Sub Combo_Click()
 If Combo.ListIndex <> 0 Then
  Dim i As Integer
  Dim r As Integer
  For i = 0 To 29
   Check1(i).Value = 0
  Next
  Randomize Timer
  i = 0
  Do
m1:
   r = Int(Rnd(1) * 29)
   If Check1(r).Value = 1 Then GoTo m1
   Check1(r).Value = 1
   Rnd_S(i) = Check1(r).Caption
   i = i + 1
  Loop While i <> Combo.ListIndex
  Rnd_S_Max = i - 1
  If Flag_ReRnd Then Call Com_Gen_Click
 Else
  For i = 0 To 29
   Check1(i).Value = 0
  Next
 End If
End Sub

Private Sub Command1_Click()
 Clipboard.Clear
 Clipboard.SetText Text_Result.Text
End Sub

Private Sub Form_Load()
 Dim i As Integer
 For i = 0 To 29
  Check1(i).Caption = Chr(33 + i)
 Next
 Combo.Clear
 Combo.AddItem "Нет"
 Combo.AddItem "1 символ"
 Combo.AddItem "2 символа"
 Combo.AddItem "3 символа"
 Combo.AddItem "4 символа"
 Combo.AddItem "5 символов"
 Combo.AddItem "6 символов"
 Combo.AddItem "7 символов"
 Combo.AddItem "8 символов"
 Combo.AddItem "9 символов"
 Combo.AddItem "10 символов"
 Combo.AddItem "11 символов"
 Combo.AddItem "12 символов"
 Combo.AddItem "13 символов"
 Combo.AddItem "14 символов"
 Combo.AddItem "15 символов"
 Combo.ListIndex = 15
 Combo_Proc.Clear
 Combo_Proc.AddItem "10%"
 Combo_Proc.AddItem "20%"
 Combo_Proc.AddItem "30%"
 Combo_Proc.AddItem "40%"
 Combo_Proc.AddItem "50%"
 Combo_Proc.AddItem "60%"
 Combo_Proc.AddItem "70%"
 Combo_Proc.AddItem "80%"
 Combo_Proc.AddItem "90%"
 Combo_Proc.AddItem "100%"
 Combo_Proc.ListIndex = 9
End Sub
