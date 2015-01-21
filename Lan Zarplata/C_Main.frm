VERSION 5.00
Begin VB.Form C_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Зарплата (клиент)"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "C_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frame 
      Height          =   4815
      Index           =   2
      Left            =   6840
      ScaleHeight     =   4755
      ScaleWidth      =   6435
      TabIndex        =   55
      Top             =   120
      Width           =   6495
      Begin VB.Frame Frame2 
         Caption         =   "Управление пользователями"
         Height          =   4095
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   6255
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'Нет
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   6015
            TabIndex        =   62
            Top             =   240
            Width           =   6015
            Begin VB.ComboBox Com_List 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   120
               Width           =   2415
            End
            Begin VB.CommandButton Com_Del 
               Caption         =   "Удалить"
               Height          =   255
               Left            =   4560
               TabIndex        =   26
               Top             =   120
               Width           =   1335
            End
            Begin VB.TextBox Text_Login 
               Height          =   285
               Left            =   2040
               TabIndex        =   27
               Top             =   960
               Width           =   3855
            End
            Begin VB.TextBox Text_Pass 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   2040
               PasswordChar    =   "*"
               TabIndex        =   28
               Top             =   1440
               Width           =   3855
            End
            Begin VB.TextBox Text_Pass2 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   2040
               PasswordChar    =   "*"
               TabIndex        =   29
               Top             =   1920
               Width           =   3855
            End
            Begin VB.TextBox Text_Path 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   2280
               TabIndex        =   30
               Top             =   2400
               Width           =   3615
            End
            Begin VB.CheckBox Check_Edit_Time 
               Caption         =   "Разрешить вводить время вручную"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2880
               Width           =   5775
            End
            Begin VB.CheckBox Check_Enter_Obed 
               Caption         =   "Разрешить вводить время обеда"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   3360
               Width           =   5775
            End
            Begin VB.Label Label8 
               Caption         =   "Пользователь:"
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
               TabIndex        =   67
               Top             =   120
               Width           =   1815
            End
            Begin VB.Line Line2 
               BorderWidth     =   2
               X1              =   120
               X2              =   5880
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label Label9 
               Caption         =   "Логин (Имя):"
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
               TabIndex        =   66
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label10 
               Caption         =   "Пароль:"
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
               TabIndex        =   65
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label20 
               Caption         =   "Пароль (еще раз):"
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
               TabIndex        =   64
               Top             =   1920
               Width           =   1815
            End
            Begin VB.Label Label21 
               Caption         =   "Путь до базы: ./base/"
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
               TabIndex        =   63
               Top             =   2400
               Width           =   2055
            End
         End
      End
      Begin VB.CommandButton Com_Edit_Confirm 
         Caption         =   "Применить изменения"
         Height          =   375
         Left            =   4080
         TabIndex        =   35
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CommandButton Com_AddNew 
         Caption         =   "Добавить как нового"
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CommandButton Com_ReportAdmin 
         Caption         =   "Отчет"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   4320
         Width           =   1335
      End
   End
   Begin VB.PictureBox Frame 
      Height          =   4815
      Index           =   1
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   6435
      TabIndex        =   52
      Top             =   5160
      Width           =   6495
      Begin VB.ComboBox Com_Buh_List 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ставка в час для работника"
         Height          =   2775
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   6255
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'Нет
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   6015
            TabIndex        =   68
            Top             =   240
            Width           =   6015
            Begin VB.TextBox Text_Buh_Stav 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1680
               TabIndex        =   19
               Top             =   1440
               Width           =   2055
            End
            Begin VB.CommandButton Com_Stav 
               Caption         =   "Изменить"
               Height          =   375
               Left            =   2280
               TabIndex        =   20
               Top             =   2040
               Width           =   1455
            End
            Begin VB.OptionButton Opt_Stav 
               Caption         =   "Ставка расчитывается по формуле (ставка=оклад/раб.дней/раб.часы)"
               Height          =   195
               Index           =   0
               Left            =   0
               MaskColor       =   &H8000000F&
               TabIndex        =   16
               Top             =   120
               Value           =   -1  'True
               Width           =   6015
            End
            Begin VB.OptionButton Opt_Stav 
               Caption         =   "Ставка в час назначается"
               Height          =   195
               Index           =   1
               Left            =   0
               MaskColor       =   &H8000000F&
               TabIndex        =   18
               Top             =   1080
               Width           =   6015
            End
            Begin VB.TextBox Text_Buh_ZP 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1680
               TabIndex        =   17
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label27 
               Caption         =   "Ставка в час.:"
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
               TabIndex        =   70
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Оклад:"
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
               TabIndex        =   69
               Top             =   600
               Width           =   735
            End
         End
      End
      Begin VB.CommandButton Com_Report 
         Caption         =   "Отчет"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton Com_DontWork 
         Caption         =   "Не рабочие дни"
         Height          =   375
         Left            =   4800
         TabIndex        =   24
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton Com_Sett 
         Caption         =   "Настройки"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Отчет по всем"
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Работник:"
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
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.PictureBox Frame 
      Height          =   4815
      Index           =   0
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   6435
      TabIndex        =   36
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Com_Send 
         Caption         =   "Отправить"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   6255
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'Нет
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   6015
            TabIndex        =   57
            Top             =   240
            Width           =   6015
            Begin VB.CommandButton Com_ReportUser 
               Caption         =   "Отчет"
               Height          =   375
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   1215
            End
            Begin VB.CommandButton Com_Vihodnie 
               Caption         =   "Выходные"
               Height          =   375
               Left            =   1320
               TabIndex        =   14
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Оклад:"
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
               Left            =   3000
               TabIndex        =   61
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label_User_ZP 
               Height          =   255
               Left            =   4560
               TabIndex        =   60
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label11 
               Caption         =   "Ставка в час.:"
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
               TabIndex        =   59
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label_Stav 
               Height          =   255
               Left            =   1680
               TabIndex        =   58
               Top             =   600
               Width           =   1215
            End
         End
      End
      Begin VB.CommandButton Com_Now 
         Caption         =   "Текущее время"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Com_Now 
         Caption         =   "Текущее время"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Com_Now 
         Caption         =   "Текущее время"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Com_Now 
         Caption         =   "Текущее время"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   11
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text_Hr 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "00"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text_Min 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "00"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text_Hr 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text_Hr 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text_Hr 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "00"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text_Min 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text_Min 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text_Min 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "00"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Время прихода:"
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
         TabIndex        =   51
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Начало обеда:"
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
         TabIndex        =   50
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Конец обеда:"
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
         TabIndex        =   49
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Время ухода:"
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
         TabIndex        =   48
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Время и дата:"
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
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label_Time_Now 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   46
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label12 
         Caption         =   "час."
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
         Left            =   2520
         TabIndex        =   45
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "час."
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
         Left            =   2520
         TabIndex        =   44
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "час."
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
         Left            =   2520
         TabIndex        =   43
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "час."
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
         Left            =   2520
         TabIndex        =   42
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "мин."
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
         Left            =   3720
         TabIndex        =   41
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "мин."
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
         Left            =   3720
         TabIndex        =   40
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "мин."
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
         Left            =   3720
         TabIndex        =   39
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "мин."
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
         Left            =   3720
         TabIndex        =   38
         Top             =   2520
         Width           =   375
      End
   End
End
Attribute VB_Name = "C_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim User_Login As String '  Логин
Dim Edit_Time As Boolean '  Разрешение вводить время вручную
Dim Time_Now As Long '      Текущее время и дата
Dim Users_List As String '  Список пользователей
Dim Last_User As Long '     Последний выбранный пользователь

Private Sub Com_AddNew_Click()
 Com_AddNew.Enabled = False
 Dim e As String
 Dim a As String
 Dim Temp() As String
 Dim i As Long
 Dim Flag As Boolean
 If Text_Login.Text = vbNullString Then
  e = e & "Не введено имя пользователя" & vbCrLf
 Else
  Text_Login.Text = Trim(Text_Login.Text)
  Temp() = Split(Users_List, vbCrLf)
  For i = 0 To UBound(Temp())
   If UCase(Temp(i)) = UCase(Text_Login.Text) Then Flag = True: Exit For
  Next
  If Flag Then e = e & "Введенное имя уже занято" & vbCrLf
 End If
 If Text_Pass.Text & Text_Pass2.Text = vbNullString Then
  e = e & "Не введен пароль" & vbCrLf
 Else
  If Text_Pass.Text <> Text_Pass2.Text Then
   e = e & "Пароли не совпадают" & vbCrLf
  End If
 End If
 If Text_Path.Text = vbNullString Then
  e = e & "Не введен путь до базы" & vbCrLf
 End If
 If e <> vbNullString Then
  MsgBox "Обнаружены ошибки:" & vbCrLf & vbCrLf & e, vbCritical, "Ошибки"
  Call Com_List_Click
 Else
  a = SendCom("Add_User|" & Text_Login.Text & "|" & Text_Pass.Text & "|" & Check_Edit_Time.Value & "|" & Check_Enter_Obed.Value & "|" & Text_Path.Text)
  If a = "Ok" Then
   Call Update_List
   Com_List.ListIndex = Com_List.ListCount - 1
  Else
   MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
  End If
 End If
 Com_AddNew.Enabled = True
End Sub

Private Sub Com_Buh_List_Click()
 Com_Buh_List.Enabled = False
 Last_User = Com_List.ListIndex
 Dim a As String
 Dim Temp() As String
 a = SendCom("Data_For_User " & Com_Buh_List.List(Com_Buh_List.ListIndex))
 Temp() = Split(a, "|")
 Opt_Stav(0).Value = IIf(Temp(0) = "1", True, False)
 Opt_Stav(1).Value = IIf(Temp(0) = "0", True, False)
 Text_Buh_ZP.Text = Temp(1)
 Text_Buh_Stav.Text = Temp(2)
 Com_Buh_List.Enabled = True
End Sub

Private Sub Com_Del_Click()
 Com_Del.Enabled = False
 Dim Log As String
 Dim a As String
 Log = Com_List.List(Com_List.ListIndex)
 If MsgBox("Вы действительно желаете удалить пользователя " & Log & " ?", vbQuestion + vbYesNo, "Удаление пользователя") = vbYes Then
  a = SendCom("Delete_User " & Log)
  If a = "Ok" Then
   Last_User = 0
   Call Update_List
  Else
   MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
  End If
 End If
 Com_Del.Enabled = True
End Sub

Private Sub Com_DontWork_Click()
 Form_Buh = True
 DontWork.Show vbModal, Me
End Sub

Private Sub Com_Edit_Confirm_Click()
 Com_Edit_Confirm.Enabled = False
 Dim a As String
 Dim e As String
 Dim Temp() As String
 Dim i As Long
 Dim Flag As Boolean
 Dim Log As String
 Log = Com_List.List(Com_List.ListIndex)
 If Text_Login.Text = vbNullString Then
  e = e & "Не введено имя пользователя" & vbCrLf
 Else
  Text_Login.Text = Trim(Text_Login.Text)
  If UCase(Log) <> UCase(Text_Login.Text) Then
   Temp() = Split(Users_List, vbCrLf)
   For i = 0 To UBound(Temp())
    If UCase(Temp(i)) = UCase(Text_Login.Text) Then Flag = True: Exit For
   Next
   If Flag Then e = e & "Введенное имя уже занято" & vbCrLf
  End If
 End If
 If Text_Pass.Text & Text_Pass2.Text = vbNullString Then
  e = e & "Не введен пароль" & vbCrLf
 Else
  If Text_Pass.Text <> Text_Pass2.Text Then
   e = e & "Пароли не совпадают" & vbCrLf
  End If
 End If
 If Text_Path.Text = vbNullString Then
  e = e & "Не введен путь до базы" & vbCrLf
 End If
 If e <> vbNullString Then
  MsgBox "Обнаружены ошибки:" & vbCrLf & vbCrLf & e, vbCritical, "Ошибки"
  Call Com_List_Click
 Else
  a = SendCom("Edit_User|" & Log & "|" & Text_Login.Text & "|" & Text_Pass.Text & "|" & Check_Edit_Time.Value & "|" & Check_Enter_Obed.Value & "|" & Text_Path.Text)
  If a = "Ok" Then
   Call Update_List
  Else
   MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
  End If
 End If
 Com_Edit_Confirm.Enabled = True
End Sub

Private Sub Com_List_Click()
 Com_List.Enabled = False
 Last_User = Com_List.ListIndex
 Dim Log As String
 Dim a As String
 Dim Temp() As String
 Log = Com_List.List(Com_List.ListIndex)
 Text_Login.Text = Log
 a = SendCom("Data_For_User " & Log)
 Temp() = Split(a, "|")
 Check_Edit_Time.Value = CLng(Temp(3))
 Check_Enter_Obed.Value = CLng(Temp(4))
 Text_Path.Text = Temp(5)
 a = SendCom("Login " & Log)
 Temp() = Split(a, "|")
 Text_Pass.Text = Temp(0)
 Text_Pass2.Text = Temp(0)
 Com_List.Enabled = True
End Sub

Private Sub Com_Report_Click()
 Form_Buh = True
 Form_User = Com_Buh_List.List(Com_Buh_List.ListIndex)
 Report.Show vbModal, Me
End Sub

Private Sub Com_ReportAdmin_Click()
 Form_Buh = False
 Form_User = Com_List.List(Com_List.ListIndex)
 Report.Show vbModal, Me
End Sub

Private Sub Com_ReportUser_Click()
 Form_Buh = False
 Form_User = User_Login
 Report.Show vbModal, Me
End Sub

Private Sub Com_Sett_Click()
 Sett_Buh.Show vbModal, Me
End Sub

Private Sub Com_Stav_Click()
 Com_Stav.Enabled = False
 Dim a As String
 a = SendCom("Change_Stav|" & Com_Buh_List.List(Com_Buh_List.ListIndex) & "|" & IIf(Opt_Stav(0).Value, "1", "0") & "|" & Text_Buh_ZP.Text & "|" & Text_Buh_Stav.Text)
 If a <> "Ok" Then MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 Com_Stav.Enabled = True
End Sub

Private Sub Com_Now_Click(Index As Integer)
 Com_Now(Index).Enabled = False
 Dim H As Long ' Часы
 Dim m As Long ' Минуты
 Dim t As String
 Time_Now = SendCom("Get_Time_Now")
 t = Time_Unix(Time_Now)
 Label_Time_Now.Caption = t
 H = CLng(Mid$(t, 1, 2))
 m = CLng(Mid$(t, 4, 2))
 Text_Hr(Index).Text = IIf(H < 10, "0", "") & H
 Text_Min(Index).Text = IIf(m < 10, "0", "") & m
 Com_Now(Index).Enabled = True
End Sub

Private Sub Com_Send_Click()
 Com_Send.Enabled = False
 Dim a As String ' Ответ
 Dim t As String ' Текущее время
 Dim s As String ' Строка запроса
 Dim D As Long '   День
 Dim m As Long '   Месяц
 Dim y As Long '   Год
 Dim i As Long
 Dim s2 As Long
 Dim s3 As Long
 Dim Flag As Boolean
 Time_Now = SendCom("Get_Time_Now")
 t = Time_Unix(Time_Now)
 Label_Time_Now.Caption = t
 D = CLng(Mid$(t, 7, 2))
 m = CLng(Mid$(t, 10, 2))
 y = CLng(Mid$(t, 13, 4))
 For i = 0 To 3
  s2 = Trim(Unix_Time(CLng(Text_Hr(i).Text), CLng(Text_Min(i).Text), D, m, y))
  s = s & "," & s2
  If i > 0 And s2 <= s3 Then Flag = True
  s3 = s2
  If Com_Now(i).Enabled Then Exit For
 Next
 If i = 0 Then s = s & ",0,0,0"
 If i = 1 Then s = s & ",0,0"
 If i = 2 Then s = s & ",0"
 If Flag Then
  MsgBox "Одно из указанных значений меньше или равно предыдущему.", vbCritical, "Ошибка"
 Else
  a = SendCom("Record|" & Trim(SID) & "|" & Mid$(s, 2))
  If a = "Ok" Then Unload Me Else MsgBox "Ошибка отправки! Попробуйте еще раз.", vbCritical, "Ошибка"
 End If
 Com_Send.Enabled = True
End Sub

Private Sub Com_Vihodnie_Click()
 Form_Buh = False
 DontWork.Show vbModal, Me
End Sub

Private Sub Command1_Click()
 Report_All.Show vbModal, Me
End Sub

Private Sub Form_Load()
 ' User_Type = 0-Пользователь, 1-бухгалтер, 2-администратор
 Dim i As Long
 Dim a As String
 Dim Temp() As String
 For i = 0 To 2
  Frame(i).BorderStyle = 0
  Frame(i).Visible = False
 Next
 Frame(User_Type).Left = 120
 Frame(User_Type).Top = 120
 Frame(User_Type).Visible = True
 ' Сбор первоначальных данных
 Select Case User_Type
  Case 0 ' Пользователь
   Me.Caption = "Зарплата (Пользователь)"
   Time_Now = SendCom("Get_Time_Now")
   Label_Time_Now.Caption = Time_Unix(Time_Now)
   User_Login = SendCom("Get_Login " & SID)
   a = SendCom("Data_For_User " & User_Login)
   Temp() = Split(a, "|")
   Label_User_ZP.Caption = Temp(1)
   Label_Stav.Caption = Temp(2)
   Edit_Time = IIf(Temp(3) = "1", True, False)
   Call Put_Tek_Dat(SendCom("Get_Data " & SID))
  Case 1 ' Бухгалтер
   Me.Caption = "Зарплата (Бухгалтер)"
   Call Update_List
  Case 2 ' Администратор
   Me.Caption = "Зарплата (Администратор)"
   Call Update_List
 End Select
End Sub

Private Sub Put_Tek_Dat(Dat As String)
 Dim i As Long
 Dim t As String ' Время
 Dim H As Long '   Часы
 Dim m As Long '   Минуты
 Dim Temp() As String
 Temp() = Split(Dat, ",")
 For i = 0 To 3
  Text_Hr(i).Enabled = False
  Text_Min(i).Enabled = False
  Com_Now(i).Enabled = False
  Com_Now(i).Visible = False
  t = Time_Unix(CLng(Temp(i)))
  H = CLng(Mid$(t, 1, 2))
  m = CLng(Mid$(t, 4, 2))
  Text_Hr(i).Text = IIf(H < 10, "0", "") & H
  Text_Min(i).Text = IIf(m < 10, "0", "") & m
 Next
 If CLng(Temp(0)) + CLng(Temp(1)) + CLng(Temp(2)) + CLng(Temp(3)) = 0 Then
  If Edit_Time Then
   Text_Hr(0).Enabled = True
   Text_Min(0).Enabled = True
  End If
  Com_Now(0).Enabled = True
  Com_Now(0).Visible = True
  Call Com_Now_Click(0)
  Exit Sub
 End If
 If CLng(Temp(0)) <> 0 And (CLng(Temp(1)) + CLng(Temp(2)) + CLng(Temp(3)) = 0) Then
  If Edit_Time Then
   Text_Hr(1).Enabled = True
   Text_Min(1).Enabled = True
  End If
  Com_Now(1).Enabled = True
  Com_Now(1).Visible = True
  Call Com_Now_Click(1)
  Exit Sub
 End If
 If CLng(Temp(0)) <> 0 And CLng(Temp(1)) <> 0 And (CLng(Temp(2)) + CLng(Temp(3)) = 0) Then
  If Edit_Time Then
   Text_Hr(2).Enabled = True
   Text_Min(2).Enabled = True
  End If
  Com_Now(2).Enabled = True
  Com_Now(2).Visible = True
  Call Com_Now_Click(2)
  Exit Sub
 End If
 If CLng(Temp(0)) <> 0 And CLng(Temp(1)) <> 0 And CLng(Temp(2)) <> 0 And CLng(Temp(3)) = 0 Then
  If Edit_Time Then
   Text_Hr(3).Enabled = True
   Text_Min(3).Enabled = True
  End If
  Com_Now(3).Enabled = True
  Com_Now(3).Visible = True
  Call Com_Now_Click(3)
  Exit Sub
 End If
 Com_Send.Enabled = False
End Sub

Private Sub Update_List()
 Dim i As Long
 Dim Tmp As Long
 Dim Temp() As String
 Users_List = SendCom("Get_Users_List")
 If Users_List = "Error" Then
  Add_Status "Ошибка при получении списка пользователей"
 Else
  Com_List.Clear
  Com_Buh_List.Clear
  Temp() = Split(Users_List, vbCrLf)
  For i = 0 To UBound(Temp()) - 3
   Com_List.AddItem Temp(i)
   Com_Buh_List.AddItem Temp(i)
  Next
  If Com_List.ListCount <> -1 Then
   If Com_List.ListCount >= Last_User Then
    Tmp = Last_User
    Com_Buh_List.ListIndex = Tmp
    Com_List.ListIndex = Tmp
   Else
    Com_Buh_List.ListIndex = 0
    Com_List.ListIndex = 0
   End If
  End If
 End If
End Sub

Private Sub Text_Buh_Stav_GotFocus()
 Text_Buh_Stav.SelStart = 0
 Text_Buh_Stav.SelLength = Len(Text_Buh_Stav.Text)
End Sub

Private Sub Text_Buh_Stav_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 Or KeyAscii <> 44) Then KeyAscii = 0
End Sub

Private Sub Text_Buh_ZP_GotFocus()
 Text_Buh_ZP.SelStart = 0
 Text_Buh_ZP.SelLength = Len(Text_Buh_ZP.Text)
End Sub

Private Sub Text_Buh_ZP_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 Or KeyAscii <> 44) Then KeyAscii = 0
End Sub

Private Sub Text_Hr_GotFocus(Index As Integer)
 Text_Hr(Index).SelStart = 0
 Text_Hr(Index).SelLength = 2
End Sub

Private Sub Text_Hr_KeyPress(Index As Integer, KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text_Hr_Validate(Index As Integer, Cancel As Boolean)
 If Val(Text_Hr(Index)) = 0 Or (Val(Text_Hr(Index)) < 0 Or Val(Text_Hr(Index)) > 23) Then Text_Hr(Index).Text = "00"
End Sub

Private Sub Text_Min_GotFocus(Index As Integer)
 Text_Min(Index).SelStart = 0
 Text_Min(Index).SelLength = 2
End Sub

Private Sub Text_Min_KeyPress(Index As Integer, KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text_Min_Validate(Index As Integer, Cancel As Boolean)
 If Val(Text_Min(Index)) = 0 Or (Val(Text_Min(Index)) < 0 Or Val(Text_Min(Index)) > 59) Then Text_Min(Index).Text = "00"
End Sub
