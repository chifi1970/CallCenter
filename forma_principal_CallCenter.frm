VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form forma_main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Appointment Scheduler"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18105
   Icon            =   "forma_principal_CallCenter.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btncarga_oficina 
      Caption         =   "..."
      Height          =   255
      Left            =   13080
      TabIndex        =   131
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton op_search 
      BackColor       =   &H00000000&
      Caption         =   "Office"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   12360
      TabIndex        =   130
      Top             =   1320
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   15000
      Sorted          =   -1  'True
      TabIndex        =   120
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   15000
      Pattern         =   "*).tt2x"
      TabIndex        =   119
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Project1.ctlCalendar Calendar2 
      Height          =   2325
      Left            =   11520
      TabIndex        =   109
      Top             =   5400
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   4101
      ShowLastMonthButton=   -1  'True
      ShowNextMonthButton=   -1  'True
      ShowLastMonthDays=   -1  'True
      ShowNextMonthDays=   -1  'True
      ShowTodayLabel  =   -1  'True
      ColorBackgroundHeader=   7526641
      ColorForegroundHeader=   16777215
      ColorSelectedBack=   0
      ColorSelectedFore=   16777215
      ColorToday      =   7526641
      ColorDayColumn  =   8388608
      ColorAlarms     =   0
      ColorBackground =   16777215
      ColorForeground =   0
      ColorButtons    =   -2147483633
      ColorLastNextMonthDayColor=   8421504
      ColorLine       =   -2147483640
      ColorWeekNumber =   8421504
      WeekStartsWith  =   1
      ShowSelected    =   -1  'True
      ShowToolTipText =   -1  'True
      ShowWeekNumbers =   0   'False
      ShowWeekNumberLeft=   -1  'True
      AllowRightClick =   0   'False
      UseAlarms       =   0   'False
      ShowShortDays   =   0   'False
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontToday {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontColumn {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ctlCalendar Calendar3 
      Height          =   2325
      Left            =   11520
      TabIndex        =   108
      Top             =   5400
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   4101
      ShowLastMonthButton=   -1  'True
      ShowNextMonthButton=   -1  'True
      ShowLastMonthDays=   -1  'True
      ShowNextMonthDays=   -1  'True
      ShowTodayLabel  =   -1  'True
      ColorBackgroundHeader=   33023
      ColorForegroundHeader=   16777215
      ColorSelectedBack=   0
      ColorSelectedFore=   16777215
      ColorToday      =   255
      ColorDayColumn  =   8388608
      ColorAlarms     =   0
      ColorBackground =   16777215
      ColorForeground =   0
      ColorButtons    =   -2147483633
      ColorLastNextMonthDayColor=   8421504
      ColorLine       =   -2147483640
      ColorWeekNumber =   8421504
      WeekStartsWith  =   1
      ShowSelected    =   -1  'True
      ShowToolTipText =   -1  'True
      ShowWeekNumbers =   0   'False
      ShowWeekNumberLeft=   -1  'True
      AllowRightClick =   0   'False
      UseAlarms       =   0   'False
      ShowShortDays   =   0   'False
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontToday {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontColumn {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.lvButtons_H btnbusca 
      Height          =   735
      Left            =   11520
      TabIndex        =   110
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":3336E
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnshow_dates 
      Height          =   735
      Left            =   11520
      TabIndex        =   112
      Top             =   3960
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1296
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":34CD5
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnsort 
      Height          =   660
      Left            =   11560
      TabIndex        =   111
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1164
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483637
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":36C5B
      ImgSize         =   40
      cBack           =   16777215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15600
      Top             =   3240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   12000
      TabIndex        =   71
      Top             =   3720
      Width           =   2055
      Begin VB.OptionButton op_showdate 
         BackColor       =   &H00000000&
         Caption         =   "Tomorrow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   73
         Top             =   580
         Width           =   1095
      End
      Begin VB.OptionButton op_showdate 
         BackColor       =   &H00000000&
         Caption         =   "Today"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   116
         Top             =   280
         Width           =   735
      End
      Begin VB.TextBox txtdate2 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   640
         TabIndex        =   78
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtdate1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   640
         TabIndex        =   76
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton op_showdate 
         BackColor       =   &H00000000&
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton op_showdate 
         BackColor       =   &H00000000&
         Caption         =   "Yesterday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1000
         TabIndex        =   72
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Appointments "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   127
         Top             =   45
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   315
         TabIndex        =   77
         Top             =   1340
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   285
         TabIndex        =   75
         Top             =   1120
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         Height          =   1545
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   12000
      TabIndex        =   65
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   118
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Office"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   70
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   69
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   68
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   67
         Top             =   480
         Width           =   820
      End
      Begin VB.OptionButton op_sort 
         BackColor       =   &H00000000&
         Caption         =   "Quote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sort by"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   126
         Top             =   30
         Width           =   525
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         Height          =   1620
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.OptionButton op_search 
      BackColor       =   &H00000000&
      Caption         =   "CC Quote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   13320
      TabIndex        =   62
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton op_search 
      BackColor       =   &H00000000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   12360
      TabIndex        =   63
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton op_search 
      BackColor       =   &H00000000&
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   13320
      TabIndex        =   61
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton op_search 
      BackColor       =   &H00000000&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   12360
      TabIndex        =   60
      Top             =   840
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtbusca 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   59
      Top             =   360
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   15120
      TabIndex        =   57
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"forma_principal_CallCenter.frx":371BE
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   15120
      TabIndex        =   56
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15120
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C8F0FF&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   11355
      TabIndex        =   25
      Top             =   1800
      Width           =   11415
      Begin Project1.lvButtons_H btnsave1 
         Height          =   855
         Left            =   10440
         TabIndex        =   104
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":37249
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnstatus 
         Height          =   345
         Left            =   5640
         TabIndex        =   122
         Top             =   4020
         Visible         =   0   'False
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   609
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":38F4C
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cbostatus_gral2 
         BackColor       =   &H00FFFBFE&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   4040
         Width           =   1935
      End
      Begin Project1.lvButtons_H btnexportar 
         Height          =   540
         Left            =   9360
         TabIndex        =   105
         Top             =   1200
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   953
         Caption         =   "Import"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":394E1
         ImgSize         =   24
         cBack           =   16777215
      End
      Begin VB.CheckBox chkimport 
         BackColor       =   &H00C8F0FF&
         Caption         =   "Automatic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   117
         Top             =   1560
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk_agente 
         BackColor       =   &H00C8F0FF&
         Caption         =   "Show only by:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   115
         Top             =   180
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin Project1.lvButtons_H btntranfiere 
         Height          =   375
         Left            =   2000
         TabIndex        =   114
         Top             =   1560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3A1D3
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.ctlCalendar Calendar1 
         Height          =   2325
         Left            =   240
         TabIndex        =   107
         Top             =   3480
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   4101
         ShowLastMonthButton=   -1  'True
         ShowNextMonthButton=   -1  'True
         ShowLastMonthDays=   -1  'True
         ShowNextMonthDays=   -1  'True
         ShowTodayLabel  =   -1  'True
         ColorBackgroundHeader=   65535
         ColorForegroundHeader=   0
         ColorSelectedBack=   4210752
         ColorSelectedFore=   16777215
         ColorToday      =   7526641
         ColorDayColumn  =   8388608
         ColorAlarms     =   0
         ColorBackground =   16777215
         ColorForeground =   0
         ColorButtons    =   -2147483633
         ColorLastNextMonthDayColor=   8421504
         ColorLine       =   -2147483640
         ColorWeekNumber =   8421504
         WeekStartsWith  =   1
         ShowSelected    =   -1  'True
         ShowToolTipText =   -1  'True
         ShowWeekNumbers =   0   'False
         ShowWeekNumberLeft=   -1  'True
         AllowRightClick =   0   'False
         UseAlarms       =   0   'False
         ShowShortDays   =   0   'False
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontToday {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontColumn {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox mensaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3720
         ScaleHeight     =   1065
         ScaleWidth      =   5265
         TabIndex        =   86
         Top             =   4560
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Image Image6 
            Height          =   975
            Left            =   -120
            Picture         =   "forma_principal_CallCenter.frx":3A625
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Please, wait a moment..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   555
            Width           =   4935
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Loading information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   87
            Top             =   120
            Width           =   5175
         End
      End
      Begin Project1.lvButtons_H btnlimpia_status 
         Height          =   180
         Left            =   3360
         TabIndex        =   84
         Top             =   4080
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3B03E
         ImgSize         =   40
         cBack           =   16777215
      End
      Begin VB.ComboBox cbocarrier 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtcelular 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin Project1.lvButtons_H btncarga_registros 
         Height          =   480
         Left            =   240
         TabIndex        =   64
         Top             =   4440
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   847
         Caption         =   "Created Today"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   40
         cBack           =   8421504
      End
      Begin VB.TextBox txtcomentario2 
         BackColor       =   &H00D7FBFD&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   58
         Top             =   4080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   48
         Top             =   4440
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   6025966
         BackColorBkg    =   13168895
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtCSR1 
         BackColor       =   &H00DEF8FE&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtcomentarios1 
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   24
         Top             =   3720
         Width           =   4935
      End
      Begin VB.ComboBox cbostatus_gral1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Text            =   "cbostatus_gral1"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txthwks1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   22
         Top             =   3720
         Width           =   760
      End
      Begin VB.TextBox txtvendor1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   21
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtrecibo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   15
         TabIndex        =   20
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtcp1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   12
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboestado1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtciudad1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtdireccion1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   80
         TabIndex        =   9
         Top             =   2280
         Width           =   4095
      End
      Begin VB.ComboBox cbo_time1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txttelefono1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   680
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtAfi1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtcliente1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         MaxLength       =   80
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cbo_Oficina1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Office where the appointment will be"
         Top             =   840
         Width           =   2535
      End
      Begin RichTextLib.RichTextBox txtturbo_quote 
         Height          =   375
         Left            =   7005
         TabIndex        =   83
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   6673406
         MultiLine       =   0   'False
         MaxLength       =   12
         TextRTF         =   $"forma_principal_CallCenter.frx":3B490
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.lvButtons_H btnlimpia_status2 
         Height          =   180
         Left            =   4680
         TabIndex        =   85
         Top             =   1365
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3B514
         ImgSize         =   40
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H btncargasql 
         Height          =   495
         Left            =   2160
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "X"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnnew1 
         Height          =   855
         Left            =   10440
         TabIndex        =   103
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3B966
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnsms 
         Height          =   555
         Left            =   6360
         TabIndex        =   106
         Top             =   1365
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   979
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3DCE6
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnlimpia_status3 
         Height          =   180
         Left            =   3360
         TabIndex        =   123
         Top             =   3765
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3FA12
         ImgSize         =   40
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H btnlink_carrier 
         Height          =   300
         Left            =   5040
         TabIndex        =   129
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "Get carrier"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnpassword 
         Height          =   975
         Left            =   10440
         TabIndex        =   132
         Top             =   2640
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_principal_CallCenter.frx":3FE64
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "@justautoins.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9800
         TabIndex        =   133
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   125
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   124
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment 1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Index           =   25
         Left            =   5400
         TabIndex        =   95
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label lbltime_appointment1 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   94
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblappointment1 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5160
         TabIndex        =   93
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TurboRater Quote: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   24
         Left            =   7040
         TabIndex        =   82
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Carriers:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   4080
         TabIndex        =   81
         Top             =   1365
         Width           =   615
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   5  'Dash-Dot-Dot
         FillColor       =   &H00FBDBEE&
         Height          =   735
         Left            =   2325
         Shape           =   4  'Rounded Rectangle
         Top             =   1275
         Width           =   4620
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   2320
         Picture         =   "forma_principal_CallCenter.frx":41C2E
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   285
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   240
         Picture         =   "forma_principal_CallCenter.frx":425F0
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell phone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   23
         Left            =   2640
         TabIndex        =   80
         Top             =   1320
         Width           =   810
      End
      Begin VB.Image op_cita 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "forma_principal_CallCenter.frx":43170
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   480
      End
      Begin VB.Image happy_face 
         Height          =   495
         Left            =   5640
         Picture         =   "forma_principal_CallCenter.frx":43CBA
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblagente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8640
         TabIndex        =   55
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   22
         Left            =   8020
         TabIndex        =   54
         Top             =   200
         Width           =   555
      End
      Begin VB.Label lbltime_appointment3 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         TabIndex        =   19
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lbltime_appointment2 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         TabIndex        =   17
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment 3:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Index           =   21
         Left            =   8880
         TabIndex        =   53
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label lblappointment3 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         TabIndex        =   18
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image btncalendar1 
         Height          =   465
         Left            =   1245
         Picture         =   "forma_principal_CallCenter.frx":44850
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgcita3 
         Height          =   375
         Left            =   5880
         Picture         =   "forma_principal_CallCenter.frx":453E2
         Stretch         =   -1  'True
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgcita3b 
         Height          =   375
         Left            =   5400
         Picture         =   "forma_principal_CallCenter.frx":459C4
         Stretch         =   -1  'True
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgcita2b 
         Height          =   495
         Left            =   5640
         Picture         =   "forma_principal_CallCenter.frx":46541
         Stretch         =   -1  'True
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgcita2 
         Height          =   480
         Left            =   5640
         Picture         =   "forma_principal_CallCenter.frx":470C5
         Stretch         =   -1  'True
         Top             =   6120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgcita1 
         Height          =   375
         Left            =   5880
         Picture         =   "forma_principal_CallCenter.frx":476B5
         Stretch         =   -1  'True
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgcita1b 
         Height          =   360
         Left            =   6000
         Picture         =   "forma_principal_CallCenter.frx":47C68
         Stretch         =   -1  'True
         Top             =   6240
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image op_cita 
         Height          =   480
         Index           =   2
         Left            =   4200
         Picture         =   "forma_principal_CallCenter.frx":487B2
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   480
      End
      Begin VB.Image op_cita 
         Height          =   480
         Index           =   1
         Left            =   3600
         Picture         =   "forma_principal_CallCenter.frx":48D94
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   480
      End
      Begin VB.Label lbloficina 
         Alignment       =   2  'Center
         BackColor       =   &H00F3F3F3&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         TabIndex        =   49
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   20
         Left            =   8520
         TabIndex        =   47
         Top             =   2040
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   19
         Left            =   6240
         TabIndex        =   46
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Gral:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   18
         Left            =   3720
         TabIndex        =   45
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Hwks:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   17
         Left            =   2520
         TabIndex        =   44
         Top             =   3480
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   16
         Left            =   1320
         TabIndex        =   43
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Receipt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   42
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label lblappointment2 
         Alignment       =   2  'Center
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   16
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment 2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Index           =   14
         Left            =   7200
         TabIndex        =   41
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Appointment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   13
         Left            =   3120
         TabIndex        =   40
         Top             =   2760
         Width           =   1350
      End
      Begin VB.Label lblstatus1 
         BackColor       =   &H009AE8FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   12
         Left            =   10440
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblfecha_cita1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   38
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   10
         Left            =   7680
         TabIndex        =   37
         Top             =   2040
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   9
         Left            =   6960
         TabIndex        =   36
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   8
         Left            =   4440
         TabIndex        =   35
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   33
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   680
         TabIndex        =   32
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFI Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   3
         Left            =   6360
         TabIndex        =   30
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblquote1 
         Alignment       =   2  'Center
         BackColor       =   &H0065D3FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbldate1 
         Alignment       =   2  'Center
         BackColor       =   &H0065D3FE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CC Quote #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   930
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Call Center"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   8295
      End
   End
   Begin Project1.lvButtons_H btncallcenter 
      Height          =   1335
      Left            =   240
      TabIndex        =   96
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   8421504
      Mode            =   2
      Value           =   -1  'True
      ImgAlign        =   4
      Image           =   "forma_principal_CallCenter.frx":49384
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnvendor 
      Height          =   1335
      Left            =   1560
      TabIndex        =   97
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   8421504
      Mode            =   2
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":4ADFB
      ImgSize         =   48
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnReports 
      Height          =   1335
      Left            =   2880
      TabIndex        =   98
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":4E2D3
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnemployees 
      Height          =   1335
      Left            =   4200
      TabIndex        =   99
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":5086F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnOffices 
      Height          =   1335
      Left            =   5520
      TabIndex        =   100
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":52B00
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnexit 
      Height          =   1335
      Left            =   9960
      TabIndex        =   101
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":54D88
      ImgSize         =   48
      cBack           =   14737632
   End
   Begin Project1.lvButtons_H btnexcel 
      Height          =   975
      Left            =   13560
      TabIndex        =   113
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_principal_CallCenter.frx":56F77
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   8280
      Picture         =   "forma_principal_CallCenter.frx":58D87
      Stretch         =   -1  'True
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   12000
      TabIndex        =   128
      Top             =   1080
      Width           =   285
   End
   Begin VB.Shape Shape7 
      FillStyle       =   0  'Solid
      Height          =   13095
      Left            =   11400
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   14400
      TabIndex        =   92
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   13680
      TabIndex        =   91
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   12960
      TabIndex        =   90
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12240
      TabIndex        =   89
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   9
      Left            =   14760
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   8
      Left            =   14400
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   7
      Left            =   14040
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   6
      Left            =   13680
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   5
      Left            =   13320
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   4
      Left            =   12960
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   3
      Left            =   12600
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   2
      Left            =   12240
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   1
      Left            =   11880
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   3240
      Picture         =   "forma_principal_CallCenter.frx":5AD56
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   79
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Caption         =   "dd/mm/yyyy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   52
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   51
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Hector Navarro "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   50
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   -360
      Top             =   0
      Width           =   18015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E6E6E6&
      BorderColor     =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   9120
      Width           =   4695
   End
End
Attribute VB_Name = "forma_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim matriz_oficina$(100, 4), cita As Integer, id_cita As Integer, modificada As Integer, seg2 As Integer, carrier As Integer
Dim matriz$(100, 2), seg As Integer, busqueda_activada As Integer, tipo_sort As Integer, fecha_busqueda As Integer
Dim cita_registrada As Integer

Private m_objDOMPeople As msxml2.DOMDocument60
Dim Temp As String
'Función api para desabilitar el repintado del control
Private Declare Function LockWindowUpdate Lib "user32" ( _
    ByVal hWndLock As Long) As Long
    

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)



Public Sub envia_correo()
On Error Resume Next


'nf = FreeFile
'Open fuente$ + "adjunta" For Output Shared As #nf
'a$ = fuentesys$ + "regkey.txt"
'Lock #nf
'Print #nf, a$
'Unlock #nf
'Close #nf


'  envia los archivos a los correos

asunto$ = "Reminder of your appointment with Just Auto Insurance"
    
' factura$ = "hna1970@yahoo.com.mx;hector@navaz.com.mx"
      
' ************************************************************
  ' carga el campo de direccion
   
   
     Set Rs = New ADODB.Recordset
    
    sSelect = "SELECT direccion From oficina where abreviatura='" + UCase(lbloficina.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    direccion1$ = Rs(0)
                         
    Rs.Close
          
          
          
  ' carga el campo de ciudad
   
   
     Set Rs = New ADODB.Recordset
    
    sSelect = "SELECT ciudad From oficina where abreviatura='" + UCase(lbloficina.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ciudad1$ = Rs(0)
                         
    Rs.Close
          
          
          
          ' carga el campo de estado
   
   
     Set Rs = New ADODB.Recordset
    
    sSelect = "SELECT estado From oficina where abreviatura='" + UCase(lbloficina.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    estado1$ = Rs(0)
                         
    Rs.Close
  
  
   ' carga el campo de TELEFONO
   
   
     Set Rs = New ADODB.Recordset
    
    sSelect = "SELECT telefono From oficina where abreviatura='" + UCase(lbloficina.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    telefono1$ = Rs(0)
    'telefono1$ = "(" + Left(telefono1$, 3) + ") " + Mid(telefono1$, 4, 3) + "-" + Right(telefono1$, 4)
                         
    Rs.Close
    
          
          
          ' carga el campo de zip
   
   
     Set Rs = New ADODB.Recordset
    
    sSelect = "SELECT cp From oficina where abreviatura='" + UCase(lbloficina.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    cp1$ = Rs(0)
                         
    Rs.Close
  
          
          
    direccion_completa$ = RTrim(direccion1$) + ", " + RTrim(ciudad1$) + ", " + RTrim(estado1$) + ", " + RTrim(cp1$)
          
          
          
          
     ' carga el campo de usuario


    Set Rs = New ADODB.Recordset
       
    sSelect = "SELECT nombre From employees where login='" + UCase(txtCSR1.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    name1$ = Rs(0)
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
       
    sSelect = "SELECT apellido From employees where login='" + UCase(txtCSR1.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    lastname1$ = Rs(0)
                         
    Rs.Close
    
    
    user1$ = name1$ + " " + lastname1$
          
    If user1$ <> "" Then
       user1$ = " With " + user1$
    End If
     
     
      
      '  +++++++++++++++++++++++++++++++++++++++++++++++++
      
      fuente_original$ = App.Path & "\"
      fuente$ = "c:\callcenter\"
      
      If Dir$(fuente$ + "nueva.htm") = "" Then
        FileCopy fuente_original$ + "nueva.htm", fuente$ + "nueva.htm"
      End If
      
      ' FileCopy App.Path & "\config.ini", fuente$ + "config.ini"
      
      
      Name fuente$ + "nueva.htm" As fuente$ + "nueva2.htm"

      nf2 = FreeFile
      Open fuente$ + "nueva.htm" For Output Shared As #nf2


      msg2$ = "</p><style type=" + Chr$(34) + "text/css" + Chr$(34) + "> <!--.Estilo1 {font-family: " + Chr$(34) + "Courier New" + Chr$(34) + "}--></style><span class=" + Chr$(34) + "Estilo1" + Chr$(34) + ">"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2


      'msg2$ = "</p> JUST AUTO INS: Appt reminder for " + txtcliente1.Text + "</p>"

      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
 
 
 '     msg2$ = "</p>" + "   " + "</p>"

  '    Lock #nf2
  '    Print #nf2, msg2$
  '    Unlock #nf2
 

      fecha_cita$ = Format(lblfecha_cita1.Caption, "mmmm dd")
      
      
      hora$ = Left(cbo_time1.List(cbo_time1.ListIndex), 5) + Right(cbo_time1.List(cbo_time1.ListIndex), 2)
      
           
      
      
      'If UCase(lbloficina.Caption) = "COM" Or UCase(lbloficina.Caption) = "PHO" Then
         msg2$ = "</p> JUST AUTO INS: Appt reminder for " + txtcliente1.Text + " on " + fecha_cita$ + " at " + Left(Format(hora$, "@@@@@@@"), 7) + ". "
      'Else
      '   msg2$ = "</p> Your appointment will be on " + fecha_cita$ + " at " + cbo_time1.List(cbo_time1.ListIndex) + ", in our office located at: </p>"
      'End If
      

      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
      
      
      
      
     
      
      
      If UCase(lbloficina.Caption) = "COM" Or UCase(lbloficina.Caption) = "PHO" Then
        msg2$ = msg2$ + "Call to " + telefono1$ + "." + user1$ + ". </p>"
      Else

         msg2$ = msg2$ + "@" + direccion_completa$ + "." + user1$ + ". </p>"
      End If
      
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      If txtturbo_quote.Text <> "" Then
         msg2$ = "</p> quote# " + txtturbo_quote.Text + "." + "</p>"

         Lock #nf2
         Print #nf2, msg2$
         Unlock #nf2
      End If
      
      
      If UCase(Left(cbo_Oficina1.List(cbo_Oficina1.ListIndex), 2)) = "JA" Then
        c$ = "Just Auto Insurance"
      Else
        c$ = "Just Auto Insurance"
      End If
            
      
      'msg2$ = "</p> We hope to see you soon ... " + c$ + "</p>"

      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
      
      
       
       Close nf2

      G = valido1

      valido1 = 1974
      Load Form1
     ' Form1.Show 1
      valido1 = G

      Kill fuente$ + "nueva.htm"
      Name fuente$ + "nueva2.htm" As fuente$ + "nueva.htm"
      
      



If transfiere$ = "NO SEND" Then
  MsgBox "The message could not be sent", 64, "ERROR detected"
Else
  MsgBox "The message was sent correctly", 64, "Attention"
End If

transfiere$ = ""

End Sub






Public Sub MuestraNodos(ByRef Nodos As msxml2.IXMLDOMNodeList)
  
  
On Error GoTo Err_Sub
      
      
    Dim TitulilloA As String
    Dim TitulilloB As String
  
    Dim oNodo As msxml2.IXMLDOMNode
      
    For Each oNodo In Nodos
          
        If oNodo.nodeType = 1 Then
  
            TitulilloA = UCase(oNodo.parentNode.nodeName)
  
                If TitulilloA <> TitulilloB Then
                    TitulilloB = TitulilloA
                    Temp = Temp & vbCrLf & _
                    UCase(oNodo.parentNode.nodeName) & vbCrLf & vbCrLf
                End If
  
        End If
  
        If oNodo.nodeType = 4 Or oNodo.nodeType = 3 Then
  
              
            Temp = Temp & oNodo.parentNode.nodeName & "=" & oNodo.nodeValue & vbCrLf
  
        End If
      
        'Si ese nodo tiene hijos (campos) se lo autopasa a la funcion
        If oNodo.hasChildNodes Then
            MuestraNodos oNodo.childNodes
        End If
    'DoEvents
    Next oNodo
      
    RichTextBox1.Text = Temp
    
    
    
    
    
      
Exit Sub
  
Err_Sub:
  
MsgBox Err.Description, vbCritical
'LockWindowUpdate 0&
  
End Sub

Private Sub Cargar_XML(Path_XML As String)
  
 On Error GoTo Err_Sub
  
    Dim objPeopleRoot As IXMLDOMElement
    Dim objPersonElement As IXMLDOMElement
    Dim ObjElement As IXMLDOMNode
    
  
    Dim X As IXMLDOMNodeList
      
    If Len(Dir(Path_XML)) = 0 Then
       MsgBox "El archivo " & Path_XML & _
               " No está en el directorio ." & vbNewLine & _
               " Compruebe la ruta", vbCritical
       Exit Sub
    End If
      
    'Seteamos la variable
   
 
    Set m_objDOMPeople = New DOMDocument60
  
    m_objDOMPeople.resolveExternals = True
  
    'Para que valide el documento xml
    m_objDOMPeople.validateOnParse = True
  
    
  
    'Carga el documento
    m_objDOMPeople.async = False
    Call m_objDOMPeople.Load(Path_XML)
  
    'Comprobamos si se carga
    If m_objDOMPeople.parseError.reason <> "" Then
        ' si hay un error muestra el fallo
        MsgBox m_objDOMPeople.parseError.reason
        Exit Sub
    End If
  
      
      
    Set objPeopleRoot = m_objDOMPeople.documentElement
      
    'nos devuelve el nombre del nodo
    Debug.Print objPeopleRoot.nodeName
    'con esto vemos el tipo de nodo
    Debug.Print objPeopleRoot.nodeType
    'nos devuelve el valor del nodo si es aplicable
    Debug.Print objPeopleRoot.nodeValue
    'Propiedad booleana que nos indica si un nodo tiene "hijos"
    Debug.Print objPeopleRoot.hasChildNodes
  
  
    Dim Index As Integer
    Dim lista As IXMLDOMNodeList
  
    RichTextBox1.Text = ""
    'LockWindowUpdate Me.hWnd
    Me.Enabled = False
    'Le pasamos el Nodo a MuestraNodos
    MuestraNodos m_objDOMPeople.childNodes
    'LockWindowUpdate 0&
    Me.Enabled = True
    
    
Exit Sub
  
Err_Sub:
      
MsgBox Err.Description
'LockWindowUpdate 0&
End Sub
  

Public Sub actualiza_registro()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
    
    ' asigna la region

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT region From oficina where abreviatura='" + lbloficina.Caption + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    region_de_oficina = Rs(0)
    
                         
    Rs.Close
    
    
    
     ' asigna el agente

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT agente From citas where quote='" + lblquote1.Caption + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    agente_registrado$ = Rs(0)
    
                         
    Rs.Close
    
    
    
    
    
    Set Rs = New ADODB.Recordset

                         
    ' Modifica el registro
    
    If cita = 0 Then
     c1$ = "1"
     c2$ = "0"
     c3$ = "0"
     
     'lblappointment2.Caption = ""
     'lblappointment3.Caption = ""
     'lbltime_appointment2.Caption = ""
     'lbltime_appointment3.Caption = ""
     
     lblappointment1.Caption = lblfecha_cita1.Caption
     lbltime_appointment1.Caption = cbo_time1.List(cbo_time1.ListIndex)
     
   ElseIf cita = 1 Then
     c1$ = "0"
     c2$ = "1"
     c3$ = "0"
     
     lblappointment2.Caption = lblfecha_cita1.Caption
     'lblappointment3.Caption = ""
     lbltime_appointment2.Caption = cbo_time1.List(cbo_time1.ListIndex)
     'lbltime_appointment3.Caption = ""
     
   ElseIf cita = 2 Then
     c1$ = "0"
     c2$ = "0"
     c3$ = "1"
     
     'lblappointment2.Caption = ""
     lblappointment3.Caption = lblfecha_cita1.Caption
    ' lbltime_appointment2.Caption = ""
     lbltime_appointment3.Caption = cbo_time1.List(cbo_time1.ListIndex)
     
   End If
   
   
   ' cambia formato de telefono
    
    t$ = txttelefono1.Text
    R$ = ""
    For Y = 1 To Len(txttelefono1.Text)
       If Asc(Mid(t$, Y, 1)) < Asc("0") Or Asc(Mid(t$, Y, 1)) > Asc("9") Then
       Else
         R$ = R$ + Mid(t$, Y, 1)
       End If
    Next Y
    t$ = R$
    
    
    
    If cbostatus_gral1.Text <> "" Then
       R$ = cbostatus_gral1.Text
       
       For Y = 0 To cbostatus_gral1.ListCount - 1
          If cbostatus_gral1.List(Y) = R$ Then
              cbostatus_gral1.ListIndex = Y
              Exit For
          End If
       Next Y
       
    
    End If
    

    
    
                         
   sSelect = "update citas set quote='" + lblquote1.Caption + "', fecha='" + lbldate1.Caption + "', oficina='" + cbo_Oficina1.List(cbo_Oficina1.ListIndex) + "', Cliente='" + UCase(txtcliente1.Text) + _
    "', AFI='" + txtAfi1.Text + "', telefono='" + t$ + "', direccion='" + UCase(txtdireccion1.Text) + "', ciudad='" + txtciudad1.Text + "', estado='" + cboestado1.Text + _
    "', cp='" + txtcp1.Text + "', fecha_cita1='" + lblfecha_cita1.Caption + "', hora_cita1='" + cbo_time1.List(cbo_time1.ListIndex) + "', status_gral='" + cbostatus_gral1.List(cbostatus_gral1.ListIndex) + _
    "', cita1='" + c1$ + "', cita2='" + c2$ + "', cita3='" + c3$ + "', fecha_cita2='" + lblappointment2.Caption + "', hora_cita2='" + lbltime_appointment2.Caption + _
    "', fecha_cita3='" + lblappointment3.Caption + "', hora_cita3='" + lbltime_appointment3.Caption + "', recibo='" + txtrecibo1.Text + "', vendor='" + txtvendor1.Text + _
    "', hwks='" + txthwks1.Text + "', comentario='" + txtcomentarios1.Text + "', CSR='" + txtCSR1.Text + _
    "', agente='" + agente_registrado$ + "', comentario2='" + txtcomentario2.Text + "', hora_24='" + Format(cbo_time1.List(cbo_time1.ListIndex), "hh:mm") + "', celular='" + txtcelular.Text + "', sms='" + _
    Format(carrier, "0#") + "', quote_turborater='" + txtturbo_quote.Text + "', region='" + Format(region_de_oficina, "00") + "', hora_citax='" + lbltime_appointment1.Caption + "', fecha_citax='" + lblappointment1.Caption + "', fecha_creacion='" + Left(lbldate1.Caption, 10) + "', status='" + cbostatus_gral2.List(cbostatus_gral2.ListIndex) + "' where idcita='" + Format(id_cita, "#####0") + "'"
    
      
      
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
    
    limpia_campos
    
    carga_registros
    
    
    
End Sub
Public Sub carga_registros()
On Error Resume Next


  
    ' Para la cadena de selección
    Dim sSelect As String
    
    fecha_hoy$ = Format(Now, "mm/dd/yyyy")
    
    
    Select Case tipo_sort
    Case 0
        ordena_por$ = "quote"
    Case 1
        ordena_por$ = "cliente"
    Case 2
        ordena_por$ = "hora_24"
    Case 3
        ordena_por$ = "fecha_cita1"
    Case 4
        ordena_por$ = "oficina"
    Case 5
        ordena_por$ = "Agente"
    
    End Select
    
    
    
    If busqueda_activada = 0 Then
      If chk_agente.value = 1 Or chk_agente.value = True Then
        
        Select Case fecha_busqueda
        Case 0
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and fecha_creacion='" + fecha_hoy$ + "' order by " + ordena_por$ + " asc"
        Case 1
              dia_ayer = Val(Format(Now, "y"))
              fecha_ayer$ = Format(dia_ayer, "mm/dd")
              fecha_ayer$ = fecha_ayer$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and fecha_cita1='" + fecha_ayer$ + "' order by " + ordena_por$ + " asc"
        Case 2
        dia_manana = Val(Format(Now, "y")) + 2
              fecha_manana$ = Format(dia_manana, "mm/dd")
              fecha_manana$ = fecha_manana$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and fecha_cita1='" + fecha_manana$ + "' order by " + ordena_por$ + " asc"
        
        Case 3
                            
              f1$ = Format(txtdate1.Text, "mm/dd/yyyy")
              f2$ = Format(txtdate2.Text, "mm/dd/yyyy")
              
              f1$ = "convert(datetime, '" + f1$ + "')"
              f2$ = "convert(datetime, '" + f2$ + "')"
              
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and fecha_cita1 between " + f1$ + " and " + f2$ + "  order by " + ordena_por$ + " asc"
        Case 4
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and fecha_cita1='" + fecha_hoy$ + "' order by " + ordena_por$ + " asc"

        End Select
        
        
        
      Else
        
        Select Case fecha_busqueda
        Case 0
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_creacion='" + fecha_hoy$ + "' order by " + ordena_por$ + " asc"
        Case 1
              dia_ayer = Val(Format(Now, "y"))
              fecha_ayer$ = Format(dia_ayer, "mm/dd")
              fecha_ayer$ = fecha_ayer$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_ayer$ + "' order by " + ordena_por$ + " asc"
        Case 2
        dia_manana = Val(Format(Now, "y")) + 2
              fecha_manana$ = Format(dia_manana, "mm/dd")
              fecha_manana$ = fecha_manana$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_manana$ + "' order by " + ordena_por$ + " asc"
        
        Case 3
                            
              f1$ = Format(txtdate1.Text, "mm/dd/yyyy")
              f2$ = Format(txtdate2.Text, "mm/dd/yyyy")
              
              f1$ = "convert(datetime, '" + f1$ + "')"
              f2$ = "convert(datetime, '" + f2$ + "')"
              
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1 between " + f1$ + " and " + f2$ + "  order by " + ordena_por$ + " asc"
        Case 4
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_hoy$ + "' order by " + ordena_por$ + " asc"
        End Select
      
      End If
   
    Else
     
       If chk_agente.value = 1 Or chk_agente.value = True Then
       
         Select Case busqueda_activada
         Case 1
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and cliente like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 2
           
           
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and telefono like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 3
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and quote='" + UCase(txtbusca.Text) + "' order by " + ordena_por$ + " asc"
           
         Case 4
             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and direccion like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         
         Case 5
             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + lblagente.Caption + "' and oficina like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         
         End Select
         
         
       
       Else
     
     
         Select Case busqueda_activada
         Case 1
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE cliente like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 2
           
           
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE telefono like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 3
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE quote='" + UCase(txtbusca.Text) + "' order by " + ordena_por$ + " asc"
           
         Case 4
             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE direccion like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         
         Case 5
             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE oficina like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         
         
         End Select
         
         
       End If
         
         busqueda_activada = 0
    
    End If
    

    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    Rs.Open sSelect, base, adOpenStatic

    ' Permitir redimensionar las columnas
    grid1.AllowUserResizing = flexResizeColumns

    ' define cuantos registros tiene la tabla
    columnas = Rs.Fields.Count
    filas = Rs.RecordCount
'Llenar las filas

   grid1.rows = filas + 1
    grid1.cols = columnas
    
    grid1.Clear

       
       
       
       For i = 0 To columnas
        grid1.TextMatrix(0, i) = Rs.Fields(i).Name
       Next i
       
      'Llenar las filas
      contador = 0
       For j = 1 To filas 'comenzamos en 1 porque el encabezado no se vuelve a llenar
          contador = contador + 1
          grid1.TextMatrix(j, 0) = Format(contador, "####0")
          For i = 1 To columnas
              If i = 3 Then
                 ' cambia la oficina por su inicial
                 
                 For Y = 0 To 100
                    If RTrim(UCase(matriz_oficina$(Y, 0))) = RTrim(UCase(Rs.Fields(i).value)) Then
                       inicial_oficina$ = RTrim(UCase(matriz_oficina$(Y, 1)))
                       Exit For
                    End If
                 Next Y
                 
                 grid1.TextMatrix(j, i) = inicial_oficina$
                 
              Else
                
                 ' =================================
                 If i = 24 Then
                     If Rs.Fields(i).value = "0" Then
                        grid1.TextMatrix(j, i) = ""
                     Else
                        grid1.TextMatrix(j, i) = Rs.Fields(i).value
                     End If
                     
                 Else
                     grid1.TextMatrix(j, i) = Rs.Fields(i).value
                 End If
                 ' =================================
                 
              End If
              
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
    
    
    ' asigna anchos de columnas
    grid1.ColWidth(0) = 900
    grid1.ColWidth(1) = 1000
    grid1.ColWidth(2) = 1800  ' agente
    grid1.ColWidth(3) = 750
    grid1.ColWidth(4) = 1750
    grid1.ColWidth(5) = 1450 ' appointment
    grid1.ColWidth(6) = 820
    grid1.ColWidth(7) = 1800
    grid1.ColWidth(8) = 3000  ' customer
    
    grid1.ColWidth(9) = 1450 ' appointment 2
    grid1.ColWidth(10) = 820
    
    
    grid1.ColWidth(11) = 1450 ' appointment 2
    grid1.ColWidth(12) = 820
    grid1.ColWidth(13) = 1450 ' appointment 3
    grid1.ColWidth(14) = 820
    grid1.ColWidth(15) = 1400   ' telefono
    grid1.ColWidth(16) = 3200  ' direccion
    grid1.ColWidth(17) = 2000  ' ciudad
    grid1.ColWidth(18) = 600  ' estado
    grid1.ColWidth(19) = 700  ' zip
    grid1.ColWidth(20) = 1200  '
    grid1.ColWidth(21) = 1800  ' vendor
    grid1.ColWidth(22) = 1200  '
   ' grid1.ColWidth(21) = 1600  ' status_gral
    grid1.ColWidth(23) = 3000  '
    grid1.ColWidth(24) = 3000  '
    grid1.ColWidth(25) = 1800  ' csr
    grid1.ColWidth(26) = 1000
    
    
    ' cambia los titulos del GRID
    grid1.Row = 0
    
    grid1.Col = 0
    grid1.Text = ""
    
    grid1.RowHeight(0) = 600
    grid1.ColAlignment(0) = 4   ' 1=izq   4=centro  7=derecha
    grid1.ColAlignment(1) = 4
    
    For Y = 2 To 24
       grid1.ColAlignment(Y) = 1
    Next Y
    
    
    
    grid1.Col = 1
    grid1.Text = "Quote"
    
    grid1.Col = 2
    grid1.Text = "Agent"
    
    grid1.Col = 3
    grid1.Text = "Office"
    
    grid1.Col = 4
    grid1.Text = "Date"
    
    grid1.Col = 5
    grid1.Text = "Appointment"
    
    grid1.Col = 6
    grid1.Text = "Time"
    
    grid1.Col = 7
    grid1.Text = "Status"
    
    grid1.Col = 8
    grid1.Text = "Customer"
    
    grid1.Col = 9
    grid1.Text = "Appointment 1"
    
    grid1.Col = 10
    grid1.Text = "Time 1"
    
    grid1.Col = 11
    grid1.Text = "Appointment 2"
    
    grid1.Col = 12
    grid1.Text = "Time 2"
    
    grid1.Col = 13
    grid1.Text = "Appointment 3"
    
    grid1.Col = 14
    grid1.Text = "Time 3"
    
    grid1.Col = 15
    grid1.Text = "Phone"
    
    grid1.Col = 16
    grid1.Text = "Address"
    
    grid1.Col = 17
    grid1.Text = "City"
    
    grid1.Col = 18
    grid1.Text = "State"
    
    grid1.Col = 19
    grid1.Text = "ZIP"
    
    grid1.Col = 20
    grid1.Text = "Receipt"
    
    grid1.Col = 21
    grid1.Text = "Vendor"
    
    grid1.Col = 22
    grid1.Text = "Hwks"
    
    'grid1.Col = 21
    'grid1.Text = "Status_gral."
    
    grid1.Col = 23
    grid1.Text = "Comments"
    
    grid1.Col = 24
    grid1.Text = "Comments"
        
    grid1.Col = 25
    grid1.Text = "CSR"
    
    grid1.Col = 26
    grid1.Text = "Quote TR"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub


Private Sub DesactivarMenu(frm As Form)
    ' Desactiva las opciones del menú del Form (esq.superior izq)
    Dim hSysmenu As Long
    hSysmenu = GetSystemMenu(frm.hWnd, 0)
    RemoveMenu hSysmenu, 6, &H400&
    RemoveMenu hSysmenu, 5, &H400&
    RemoveMenu hSysmenu, 4, &H400&
    RemoveMenu hSysmenu, 2, &H400&
    RemoveMenu hSysmenu, 1, &H400&
End Sub

Private Sub btnbusca_Click()
On Error Resume Next

If txtbusca.Text = "" Then Exit Sub


If op_search(0).value = True Then
  busqueda_activada = 1
ElseIf op_search(1).value = True Then
  busqueda_activada = 2
ElseIf op_search(2).value = True Then
  busqueda_activada = 3
ElseIf op_search(3).value = True Then
  busqueda_activada = 4
ElseIf op_search(4).value = True Then
  busqueda_activada = 5
End If



carga_registros

End Sub

Private Sub btncalendar1_Click()
On Error Resume Next


Calendar1.Visible = True
seg2 = 0
Timer2.Enabled = True

End Sub

Private Sub btncarga_oficina_Click()
On Error Resume Next
op_search(4).value = True
txtbusca.Text = cbo_Oficina1.List(cbo_Oficina1.ListIndex)

End Sub

Private Sub btncarga_registros_Click()
On Error Resume Next

op_showdate(0).value = False
op_showdate(1).value = False
op_showdate(2).value = False

busqueda_activada = 0

fecha_busqueda = 0

carga_registros

End Sub

Private Sub btncargasql_Click()
On Error Resume Next
Load Form1
Form1.Show 1

End Sub

Private Sub btnemployees_Click()
On Error Resume Next
base.Close

Load forma_employees
forma_employees.Show
Unload Me
End Sub

Private Sub btnexcel_Click()
On Error Resume Next
Dim sData As String
 
If grid1.rows = 1 Then Exit Sub
 
archivo$ = "c:\callcenter\Callcenter.xlsx"
Kill archivo$

If Dir$(archivo$) <> "" Then
  MsgBox "Please, close the file " + archivo$ + " and try it again.", 64, "Attention"
  Exit Sub
End If

mensaje.Visible = True
mensaje.Refresh



sData = "Quote" & vbTab & "Agent" & vbTab & "Office" & vbTab & "Date" & vbTab & "Appointment1" & vbTab & "Time1" & vbTab & "status" & vbTab & "Customer" & vbTab & "Appointment2" _
& vbTab & "Time2" & vbTab & "Appointment3" & vbTab & "Time3" & vbTab & "Phone" & vbTab & "Address" & vbTab & "City" & vbTab & "State" & vbTab & "Zip" _
& vbTab & "Receipt" & vbTab & "vendor" & vbTab & "hwks" & vbTab & "Comment1" & vbTab & "Comment2" & vbTab & "csr" & vbCr


  
 
  'Create a new workbook in Excel
   Dim oExcel As Object
   Dim oBook As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.add

  

   'Clipboard.Clear

   'Clipboard.SetText sData
   
   conta = 0

   'Paste the data
   'oBook.Worksheets(1).Range("A" + Format(conta, "####0")).Select
   'oBook.Worksheets(1).Paste
   'conta = conta + 1



For t = 0 To grid1.rows - 1
  grid1.Row = t
  sData = ""
  
  For Y = 1 To 26
    grid1.Col = Y
    R$ = grid1.Text
    
    sData = sData + R$ & vbTab
  Next Y
    
    
  sData = sData & vbCr
           
   
   Clipboard.Clear

   Clipboard.SetText sData
   
   conta = conta + 1
   'Paste the data
   oBook.Worksheets(1).Range("A" + Format(conta, "####0")).Select
   oBook.Worksheets(1).Paste
            
           
Next t

   
 
  
   'Save the Workbook and Quit Excel
   oBook.SaveAs archivo$
   If Err Then
     MsgBox "You have open the file " + archivo$ + ". It couldn't save anything. Close it and try again.", 16, "Attention"
     oExcel.Quit
     mensaje.Visible = False

     Exit Sub
     
   End If
   
   oExcel.Quit
   
   mensaje.Visible = False
   
   If Dir$("C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE") <> "" Then
      R$ = Shell("C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE c:\callcenter\Callcenter.xlsx", vbNormalFocus)
   Else
      R$ = Shell("C:\Program Files\Microsoft Office\Office15\EXCEL.EXE c:\callcenter\Callcenter.xlsx", vbNormalFocus)
   End If
   
   MsgBox "The file named " + archivo$ + " was created successfully", 64, "Attention"
   


End Sub

Private Sub btnexit_Click()
On Error Resume Next
base.Close
End
End Sub








Private Sub btnexportar_Click()
On Error Resume Next
limpia_campos

n$ = "c:\transfer\IntegrationData.tt2x"
'If Dir$(n$) = "" Then
If chkimport.value = 0 Then
' GoTo saltado
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = "c:\transfer"

   ' Set flags
    CommonDialog1.flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "HawkSoft CMS (*.tt2x)|*.tt2x"
    ' Display the Save dialog box
    CommonDialog1.filename = "IntegrationData.tt2x"

    CommonDialog1.ShowOpen
    If CommonDialog1.filename = "" Then Exit Sub
    
    n$ = CommonDialog1.filename
Else
    List2.Clear
    File1.Path = "c:\transfer"
    For t = 0 To File1.ListCount - 1
       pos = InStr(1, File1.List(t), "(")
       pos2 = InStr(1, File1.List(t), ")")
       longitud = pos2 - pos
       num_file = Mid(File1.List(t), pos + 1, longitud - 1)
    
       List2.AddItem Format(num_file, "00") + " " + File1.List(t)
    Next t
    n$ = List2.List(List2.ListCount - 1)
    n$ = Right(n$, Len(n$) - 3)
        
    n$ = "c:\transfer\" + n$ 'IntegrationData.tt2x
End If
'End If

If n$ = "" Then Exit Sub
saltado:
    
    
RichTextBox1.Text = ""
Temp = ""

mensaje.Visible = True
mensaje.Refresh
 
guarda_oficina$ = ""
List1.Clear
    nf = FreeFile
    Open n$ For Input Shared As #nf
    
    veces = 0
    var_xml$ = ""
    Do Until EOF(nf)
       R$ = ""
       Lock #nf
       Line Input #nf, R$
       Unlock #nf
       
       If R$ = "" Then veces = veces + 1
       If veces > 5 Then Exit Do
       
       List1.AddItem R$
       var_xml$ = var_xml$ + R$
    Loop
    
    Close #nf
    
    
    
    
    List1.RemoveItem List1.ListCount - 1
    
    tamano_inicial = 0
    For t = 0 To List1.ListCount - 1
         tamano_inicial = tamano_inicial + Len(List1.List(t))
    Next t
    
    
        
    pos = InStr(1, n$, ".")
    n2$ = Left(n$, pos) + "XML"
    n3$ = Left(n$, pos) + "tmp"
    
    
    ' Kill n2$
    
    
    
    'Array que contendrá los bytes del archivo es decir los datos
    Dim Data As Byte
  
    'Variable Para el tamaño del archivo ( luego se usa para el Redim )
    Dim fLen As Long
  
      
  
    'Abrimos el archivo en modo binario de solo lectura (Binary Lock Read)
    Open n$ For Binary Lock Read As #1
  
  
  
    'creamos un archivo para guardar los datos ( Binary Access Write )
    Open n2$ For Output Shared As #2
    Close 2
    
    Open n2$ For Binary Access Write As #2
  
    'Redimiensionamos el array al tamaño del archivo
     fLen = FileLen(n$)
  
    'ReDim Data(fLen) As Byte
    
    Dim cont As Long
    
    
    cont = 0
    For t = 1 To FileLen(n$)
      'Leemos el archivo entero y lo almacenamos en el array
      
      Get #1, , Data
      cont = cont + 1
  
  
       
      If cont > tamano_inicial + 34 Then
          'Escribimos los bytes del array anterior, en el nuevo archivo ( archivo 2 )
          Put #2, , Data
      End If
    Next t
    
    'Cerramos los dos archivos
    Close #1, #2
    
    

      Call Cargar_XML(n2$)
      
      
      nf = FreeFile
      Open n3$ For Output Shared As #nf
      Lock #nf
      Print #nf, RichTextBox1.Text
      Unlock #nf
      Close #nf
      
      
      ' carga datos del IntegrationData file
      
      nf = FreeFile
      Open n3$ For Input Shared As #nf
      
      veces = 0
      nombrechecado = 0
      nombrechecado2 = 0
      nombrechecado3 = 0
      
      tel_modo = 0
      dir_modo = 0
      existe = 0
      
      Do Until EOF(nf)
         R$ = ""
         Lock #nf
         Line Input #nf, R$
         Unlock #nf
         
         If R$ = "" Then veces = veces + 1
         If veces > 400 Then Exit Do
         
         If Left(R$, 21) = "TransactionRequestDt=" Then
           lbldate1.Caption = Mid(R$, 27, 2) + "/" + Mid(R$, 30, 2) + "/" + Mid(R$, 22, 4) + " " + Format(Mid(R$, 33, 5), "hh:mm am/pm")
         End If
         
         ' carga oficina
         ' ====================
         
         If Left(R$, 8) = "Surname=" Then
           nombrechecado = nombrechecado + 1
           If nombrechecado = 2 Then
               apellido$ = Right(R$, Len(R$) - 8)
           End If
         End If
         
         
         If Left(R$, 10) = "GivenName=" Then
           nombrechecado2 = nombrechecado2 + 1
           If nombrechecado2 = 2 Then
               nombre1$ = Right(R$, Len(R$) - 10)
           End If
         End If
         
         
         If Left(R$, 15) = "OtherGivenName=" Then
           nombrechecado3 = nombrechecado3 + 1
           If nombrechecado3 = 2 Then
               nombre2$ = Right(R$, Len(R$) - 15)
           End If
         End If
         
         
         
         If Left(R$, 23) = "CommunicationUseCd=Home" Then
           tel_modo = 1  ' casa
         End If
         
         If Left(R$, 16) = "PhoneTypeCd=Cell" Then
           tel_modo = 2  ' celular
         End If
         
         
           If Left(R$, 12) = "PhoneNumber=" And tel_modo = 1 Then
                 tel$ = Right(R$, Len(R$) - 12)
                 txttelefono1.Text = tel$
           End If
         
           If Left(R$, 12) = "PhoneNumber=" And tel_modo = 2 Then
                 celular$ = Right(R$, Len(R$) - 12)
                 txtcelular.Text = celular$
           End If
         
         
         If Left(R$, 11) = "AddrTypeCd=" Then
           dir_modo = dir_modo + 1
         End If
         
         
         If dir_modo = 1 Then
         
            ' verifica si es la direccion correcta y la oficna
            If Left(R$, 6) = "Addr1=" Then
               d$ = Right(R$, Len(R$) - 6)
               
               
               For z = 0 To 99
                 If Right(R$, Len(R$) - 6) = matriz_oficina$(z, 2) Then
                       guarda_oficina$ = matriz_oficina$(z, 0)
                   existe = 1
                   Exit For
                 End If
               Next z

            End If
            
            
            If existe = 1 Then
              ' si encontro la direccion
                For w = 0 To cbo_Oficina1.ListCount - 1
                 If RTrim(UCase(cbo_Oficina1.List(w))) = RTrim(UCase(guarda_oficina$)) Then
                   cbo_Oficina1.ListIndex = w
                   existe = 2
                   Exit For
                 End If
                Next w
         
              
            Else
            
              If Left(R$, 5) = "City=" Then
                city$ = Right(R$, Len(R$) - 5)
                            
                hayado = 0
                For w = 0 To cbo_Oficina1.ListCount - 1
                 If Right(RTrim(UCase(cbo_Oficina1.List(w))), Len(city$)) = RTrim(UCase(city$)) Then
                   If existe <> 2 Then
                      cbo_Oficina1.ListIndex = w
                   End If
                   hayado = 1
                   Exit For
                 End If
                Next w
                       
                
              End If
         
            End If
                     
         End If
         
         
         
         
         If dir_modo = 2 Then
            If Left(R$, 6) = "Addr1=" Then
               txtdireccion1.Text = Right(R$, Len(R$) - 6)
            End If
            
            If Left(R$, 5) = "City=" Then
               txtciudad1.Text = Right(R$, Len(R$) - 5)
            End If
            
            If Left(R$, 12) = "StateProvCd=" Then
               For Y = 0 To cboestado1.ListCount - 1
                   If UCase(cboestado1.List(Y)) = UCase(Right(R$, Len(R$) - 12)) Then
                       cboestado1.ListIndex = Y
                       Exit For
                   End If
               Next Y
            End If
            
            If Left(R$, 11) = "PostalCode=" Then
               txtcp1.Text = Right(R$, Len(R$) - 11)
            End If
            
         End If
         
         
         
         
      Loop
      
      If nombre2$ <> "" Then
          txtcliente1.Text = nombre1$ + " " + nombre2$ + " " + apellido$
      Else
          txtcliente1.Text = nombre1$ + " " + apellido$
      End If
      
      Close nf
         
         
      Kill n2$
      Kill n3$
       
         
      id_cita = 0
         
      
      mensaje.Visible = False
      inhabilita_campos
      
      cbo_Oficina1.Enabled = True
      txtcomentarios1.Enabled = True
      txtCSR1.Enabled = True
      txtAfi1.Enabled = True
      cbocarrier.Enabled = True
      txtturbo_quote.Enabled = True
      txtcelular.Enabled = True

End Sub

Private Sub btnlimpia_status_Click()
On Error Resume Next
If cbostatus_gral2.Enabled = True Then
   cbostatus_gral2.Text = ""
   cbostatus_gral2.ListIndex = -1
End If

End Sub

Private Sub btnlimpia_status2_Click()
On Error Resume Next
If cbocarrier.Enabled = True Then
  cbocarrier.ListIndex = -1
End If

End Sub
 


Private Sub btnlimpia_status3_Click()
On Error Resume Next

   cbostatus_gral1.Text = ""
   cbostatus_gral1.ListIndex = -1

End Sub

Private Sub btnlink_carrier_Click()
On Error Resume Next

R$ = Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  https://freecarrierlookup.com", vbNormalFocus)
Clipboard.Clear
End Sub

Private Sub btnnew1_Click()
On Error Resume Next
' carga_oficinas
limpia_campos

End Sub

Private Sub btnOffices_Click()
On Error Resume Next
base.Close

Load forma_oficinas
forma_oficinas.Show
Unload Me

End Sub



Private Sub btnpassword_Click()
On Error Resume Next
Load forma_acceso
forma_acceso.Show 1

' verifica password
' ************************************************************
  ' carga el campo de password


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT contrasena From employees where login='" + UCase(lblagente.Caption) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    password_real$ = Rs(0)
            
                         
    Rs.Close
    
    
    If password$ = password_real$ Then
       
       password$ = lblagente.Caption
       Load forma_cambia_password
       forma_cambia_password.Show 1
       
    
    
 
       
       
    Else
       MsgBox "The password is not valid. Try again or call the IT department to change it.", 64, "Attention"
       
    End If
    
    

End Sub

Private Sub btnReports_Click()
On Error Resume Next
base.Close

Load forma_graficas
forma_graficas.Show
Unload Me
End Sub


Private Sub btnsave1_Click()
On Error Resume Next
  
If cbo_Oficina1.ListIndex = -1 Then
  MsgBox "You need to select an office", 64, "Attention"
  Exit Sub
End If
  
  
  
If txtcliente1.Text = "" Then
   MsgBox "You need to type the Customer Name", 64, "Attention"
   Exit Sub
End If


If txttelefono1.Text = "" Then
   MsgBox "You need to type the phone", 64, "Attention"
   Exit Sub
End If


If txtdireccion1.Text = "" Then
   MsgBox "You need to type the address", 64, "Attention"
   Exit Sub
End If


If txtciudad1.Text = "" Then
   MsgBox "You need to type the city", 64, "Attention"
   Exit Sub
End If


If cboestado1.ListIndex = -1 And cboestado1.Text = "" Then
   MsgBox "You need to select the state", 64, "Attention"
   Exit Sub
End If


If txtcp1.Text = "" Then
   MsgBox "You need to type the zip code", 64, "Attention"
   Exit Sub
End If


If lblfecha_cita1.Caption = "" Then
   MsgBox "You need to select the appointment", 64, "Attention"
   Exit Sub
End If


If cbo_time1.ListIndex = -1 Then
   MsgBox "You need to select the time", 64, "Attention"
   Exit Sub
End If


txtCSR1.Text = lblagente.Caption

'If txtCSR1.Text = "" Then
'   MsgBox "You need to type the CSR", 64, "Attention"
'   Exit Sub
'End If



' revisa is existe el numero de id

n$ = ""
If lblquote1.Caption <> "" Then


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idcita From citas where quote='" + lblquote1.Caption + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    n$ = Rs(0)
    id_cita = Val(n$)
                         
    Rs.Close
    
    
 End If
    
    
    
 ' verifica si fecha fue modificada
 
 
 
 If txtcelular.Text = "" Then
    cbocarrier.ListIndex = -1
 End If
 
 
 
    
    existe_usuario$ = ""
 '  revisa si existe el usuario
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT activo From employees where login='" + LTrim(RTrim(txtCSR1.Text)) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    existe_usuario$ = Rs(0)
    Rs.Close
    
    If existe_usuario$ <> "" Then
      If existe_usuario$ = "N" Then
          MsgBox "The user CSR is disabled. Please contact to IT Department.", 16, "Attention"
          Exit Sub
      End If
    Else
    
       MsgBox "The user CSR doesn't exist. Please contact to IT Department.", 16, "Attention"
       Exit Sub
    End If
      
      
    ' ------------------------------------------------------------------------------
      
      
    If txtvendor1.Text <> "" Then
      existe_usuario$ = ""
   '  revisa si existe el usuario
      Set Rs = New ADODB.Recordset
  
      sSelect = "SELECT activo From employees where login='" + LTrim(RTrim(txtvendor1.Text)) + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      existe_usuario$ = Rs(0)
      Rs.Close
    
    
      If existe_usuario$ <> "" Then
        If existe_usuario$ = "N" Then
          MsgBox "The user Vendor is disabled. Please contact to IT Department.", 16, "Attention"
          Exit Sub
        End If
      Else
    
        MsgBox "The user Vendor doesn't exist. Please contact to IT Department.", 16, "Attention"
        Exit Sub
      End If
      
    End If
    
    
      
      
    
     
 If n$ = "" Or id_cita = 0 Then
   Agrega_registro
 Else
 
  
   
   '  revisa si cita fue cambiada
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT hora_cita1 From citas where quote='" + lblquote1.Caption + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    hora_registrada$ = Rs(0)
    Rs.Close
    
    
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT fecha_cita1 From citas where quote='" + lblquote1.Caption + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    fecha_registrada$ = Rs(0)
    Rs.Close
    
    
    
    
    
    'If (hora_registrada$ <> cbo_time1.List(cbo_time1.ListIndex)) Or (lblfecha_cita1.Caption <> fecha_registrada$) Then
    If (lblfecha_cita1.Caption <> fecha_registrada$) Then
   
   
        '  revisa si cita fue cambiada
        Set Rs = New ADODB.Recordset
        sSelect = "SELECT cita1 From citas where quote='" + lblquote1.Caption + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        c1 = Rs(0)
        Rs.Close
       
        Set Rs = New ADODB.Recordset
        sSelect = "SELECT cita2 From citas where quote='" + lblquote1.Caption + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        c2 = Rs(0)
        Rs.Close
   
        Set Rs = New ADODB.Recordset
        sSelect = "SELECT cita3 From citas where quote='" + lblquote1.Caption + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        c3 = Rs(0)
        Rs.Close
   
     
                
        If c1 = 1 Then
           cita = 1
        ElseIf c2 = 1 Then
           cita = 2
        ElseIf c3 = 1 Then
           cita = 3
        End If
   
   
        op_cita(0).Picture = imgcita1.Picture
        op_cita(1).Picture = imgcita2.Picture
        op_cita(2).Picture = imgcita3.Picture

        If cita = 0 Then
             op_cita(0).Picture = imgcita1b.Picture
        ElseIf cita = 1 Then
             op_cita(1).Picture = imgcita2b.Picture
        ElseIf cita = 2 Then
             op_cita(2).Picture = imgcita3b.Picture
        ElseIf cita = 3 Then
             MsgBox "You can not create a fourth appointment with this quote. Please create a new appointment.", 16, "Attention"
             Exit Sub
        End If

        
   ElseIf (hora_registrada$ = cbo_time1.List(cbo_time1.ListIndex)) And (lblfecha_cita1.Caption = fecha_registrada$) Then
        ' Exit Sub
   
   End If
   
    
    
   
   
   actualiza_registro
 End If
 
    
agrega_nuevo_registro:
If txtCSR1.Text <> "" Then
 nf = FreeFile
 Open "c:\callcenter\CSR" For Output Shared As #nf
 Lock #nf
 Print #nf, txtCSR1.Text
 Unlock #nf
 Close #nf
End If

   
End Sub

Public Sub Agrega_registro()
On Error Resume Next
  
  
 
' revisa el numero de id disponible


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idcita From citas ORDER BY idcita DESC;"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ultimo_id = Rs(0)
    
                             
    Rs.Close
                         
                         
                         

   
   
   
   
   
' asigna la region

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT region From oficina where abreviatura='" + lbloficina.Caption + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    region_de_oficina = Rs(0)
    
                         
    Rs.Close
    
    
    
    
   

   If cita = 0 Then
     c1$ = "1"
     c2$ = "0"
     c3$ = "0"
     
     'lblappointment2.Caption = ""
     'lblappointment3.Caption = ""
     'lbltime_appointment2.Caption = ""
     'lbltime_appointment3.Caption = ""
     
     lblappointment1.Caption = lblfecha_cita1.Caption
     lbltime_appointment1.Caption = cbo_time1.List(cbo_time1.ListIndex)
     
   ElseIf cita = 1 Then
     c1$ = "0"
     c2$ = "1"
     c3$ = "0"
     
     lblappointment2.Caption = lblfecha_cita1.Caption
     'lblappointment3.Caption = ""
     lbltime_appointment2.Caption = cbo_time1.List(cbo_time1.ListIndex)
     'lbltime_appointment3.Caption = ""
     
   ElseIf cita = 2 Then
     c1$ = "0"
     c2$ = "0"
     c3$ = "1"
     
     'lblappointment2.Caption = ""
     lblappointment3.Caption = lblfecha_cita1.Caption
     'lbltime_appointment2.Caption = ""
     lbltime_appointment3.Caption = cbo_time1.List(cbo_time1.ListIndex)
     
   End If
   
                         
                         
                         
    ' inserta el registro
inserta:
                         
    ' cambia formato de telefono
    
    t$ = txttelefono1.Text
    R$ = ""
    For Y = 1 To Len(txttelefono1.Text)
       If Asc(Mid(t$, Y, 1)) < Asc("0") Or Asc(Mid(t$, Y, 1)) > Asc("9") Then
       Else
         R$ = R$ + Mid(t$, Y, 1)
       End If
    Next Y
    t$ = R$
    
      
    
    
    ' asigna el ultimo quote

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT quote From citas ORDER BY quote DESC"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ultimo_quote = Rs(0)
    
                         
    Rs.Close


    If ultimo_quote = Empty Then
        Set Rs = New ADODB.Recordset
        sSelect = "SELECT quote From citas"
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
        Rs.Open sSelect, base, adOpenUnspecified
        ultimo_quote = Rs(0)
              
        Rs.Close

    End If



   lblquote1.Caption = Format(ultimo_quote + 1, "#####0")
   lbldate1.Caption = Format(Now, "mm/dd/yyyy") + " " + Format(Now, "hh:mm am/pm")


    If cbostatus_gral1.Text <> "" Then
       R$ = cbostatus_gral1.Text
       
       For Y = 0 To cbostatus_gral1.ListCount - 1
          If cbostatus_gral1.List(Y) = R$ Then
              cbostatus_gral1.ListIndex = Y
              Exit For
          End If
       Next Y
       
    
    End If
    
    

                         
    sSelect = "INSERT INTO citas (Idcita, quote, fecha, oficina, cliente, afi, telefono, direccion, ciudad, estado, cp, fecha_cita1, hora_cita1, status_gral, cita1, cita2, cita3, fecha_cita2, hora_cita2, fecha_cita3, hora_cita3, recibo, vendor, hwks, comentario, csr, agente, comentario2, hora_24, celular, sms, quote_turborater, region, hora_citax, fecha_citax, fecha_creacion, status)  VALUES ('" + _
    Format(ultimo_id + 1, "####0") + "', '" + lblquote1.Caption + "', '" + lbldate1.Caption + "', '" + cbo_Oficina1.List(cbo_Oficina1.ListIndex) + "', '" + UCase(txtcliente1.Text) + "', '" + txtAfi1.Text + _
    "', '" + t$ + "', '" + UCase(txtdireccion1.Text) + "', '" + txtciudad1.Text + "', '" + cboestado1.List(cboestado1.ListIndex) + "', '" + txtcp1.Text + "', '" + lblfecha_cita1.Caption + "', '" + cbo_time1.List(cbo_time1.ListIndex) + _
    "', '" + cbostatus_gral1.List(cbostatus_gral1.ListIndex) + "', '" + c1$ + "', '" + c2$ + "', '" + c3$ + "', '" + lblappointment2.Caption + "', '" + lbltime_appointment2.Caption + " ', '" + lblappointment3.Caption + _
    "', '" + lbltime_appointment3.Caption + "', '" + txtrecibo1.Text + "', '" + txtvendor1.Text + "', '" + txthwks1.Text + "', '" + txtcomentarios1.Text + "', '" + txtCSR1.Text + "', '" + agente + "', '" + txtcomentario2.Text + "' ,'" + Format(cbo_time1.List(cbo_time1.ListIndex), "hh:mm") + _
    "', '" + txtcelular.Text + "', '" + Format(carrier, "0#") + "', '" + txtturbo_quote.Text + "', '" + Format(region_de_oficina, "00") + "', '" + lbltime_appointment1.Caption + "', '" + lblappointment1.Caption + "', '" + Left(lbldate1.Caption, 10) + "', '" + cbostatus_gral2.List(cbostatus_gral2.ListIndex) + "')"
    
    
   
   
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
            
    limpia_campos
    
    carga_registros
    
    
End Sub


Private Sub btnsetup_Click()



End Sub

Private Sub btnshow_dates_Click()
On Error Resume Next
limpia_campos

busqueda_activada = 0


If op_showdate(0).value = True Then
  fecha_busqueda = 1
ElseIf op_showdate(1).value = True Then
  fecha_busqueda = 2
ElseIf op_showdate(2).value = True Then
  fecha_busqueda = 3
ElseIf op_showdate(3).value = True Then
  fecha_busqueda = 4
End If





If fecha_busqueda = 3 Then   ' fecha variable

    year_previous = Val(Format(Now, "yyyy")) - 1
    month_previous = Val(Format(Now, "mm")) - 1

    year_current = Val(Format(Now, "yyyy"))
    month_current = Val(Format(Now, "mm"))
    
    year1 = Val(Right(txtdate1.Text, 4))
    year2 = Val(Right(txtdate2.Text, 4))
    
    mes1 = Val(Mid$(txtdate1.Text, 4, 2))
    mes2 = Val(Mid$(txtdate2.Text, 4, 2))
    
    If year2 = 0 Then year2 = year1
    
    
    If year1 < (year_previous) Or year2 < (year_previous) Then
       MsgBox "Date is not valid" + Chr$(13) + "Use: mm/dd/yyyy", 64, "Attention"
       Exit Sub
    End If
    
    If year1 > (year_current + 2) Or year2 > (year_current + 2) Then
       MsgBox "Date is not valid" + Chr$(13) + "Use: mm/dd/yyyy", 64, "Attention"
       Exit Sub
    End If
    
    
    
    
    

End If




carga_registros

End Sub

Private Sub btnsms_Click()
On Error Resume Next
If txtcelular.Text = "" Then
   MsgBox "You can not send a text message because you do not have a cell phone number assigned", 16, "Attention"
   Exit Sub
End If


If lblquote1.Caption = "" Then
  MsgBox "You can not send a text message because you have not saved the appointment yet", 16, "Attention"
   Exit Sub
End If

If cbocarrier.ListIndex = -1 Then
  MsgBox "You can not send a text message because you have not chosen the carrier yet", 16, "Attention"
   Exit Sub
End If

If lblquote1.Caption = "" Then Exit Sub


' quita los simbolos no numericos al celular
R$ = ""
For Y = 1 To Len(txtcelular.Text)
   If Asc(Mid$(txtcelular.Text, Y, 1)) >= Asc("0") And Asc(Mid$(txtcelular.Text, Y, 1)) <= Asc("9") Then
       R$ = R$ + Mid$(txtcelular.Text, Y, 1)
   End If
   
Next Y


' valida numero
If Len(R$) < 10 Or Len(R$) > 11 Then
   MsgBox "Please, verify the cell phone number", 16, "Invalid phone number"
   Exit Sub
End If

If Len(R$) = 11 And Left(R$, 1) <> "1" Then
   MsgBox "please, check the country area", 16, "Invalid phone number"
   Exit Sub
End If


Select Case carrier
  Case 0
    a$ = "@txt.att.net"
  Case 1
    a$ = "@myboostmobile.com"
  Case 2
    a$ = "@mobile.celloneusa.com"
  Case 3
    a$ = "@cingularme.com"
  Case 4
    a$ = "@mms.cricketwireless.net"
  Case 5
    a$ = "@mymetropcs.com"
  Case 6
    a$ = "@messaging.nextel.com"
  Case 7
    a$ = "@pcsone.net"
  Case 8
    a$ = "@messaging.sprintpcs.com"
  Case 9
    a$ = "@tmomail.net"
  Case 10
    a$ = "@email.uscc.net"
  Case 11
    a$ = "@vtext.com"
  
End Select







transfiere$ = R$ + a$
     
     
envia_correo




End Sub

Private Sub btnsort_Click()
On Error Resume Next
carga_registros

End Sub

Private Sub btnstatus_Click()
On Error Resume Next
cbostatus_gral1.Text = cbostatus_gral2.List(cbostatus_gral2.ListIndex)
cbostatus_gral2.Text = ""
cbostatus_gral2.ListIndex = -1
End Sub

Private Sub btntranfiere_Click()
On Error Resume Next
If txtcelular.Text = "" Then
   txtcelular.Text = txttelefono1.Text
Else
  R$ = MsgBox("Do you want to replace the existing cell phone with this Telephone?", 4, "verifying...")
  If R$ = "7" Then Exit Sub
  txtcelular.Text = txttelefono1.Text
  
End If
End Sub

Private Sub btnvendor_Click()
On Error Resume Next
base.Close

Load forma_vendor
forma_vendor.Show
Unload Me
End Sub

Private Sub Calendar1_DateClicked(inputDate As Date)
On Error Resume Next
     
    fecha_guardada$ = lblfecha_cita1.Caption
     
    R$ = Format(inputDate, "mm/dd/yyyy")
    
    ' cita=0 primera cita
    ' cita=1 segunda cita
    ' cita=2 tercera cita
    
    f1$ = Format(R$, "y") ' fecha actual o puesta
    f2$ = Format(fecha_guardada$, "y")  ' fecha guardada originalmente
    
    ano1$ = Right(R$, 4)  ' ano actual
    ano2$ = Right(fecha_guardada$, 4)  ' ano guardado
    
    
    If Val(ano1$) < Val(ano2$) Then
          MsgBox "Cannot select a date prior to today", 16, "Attention"
          GoTo salida
        
    End If
    
    
    
    
    If lblfecha_cita1.Caption = "" Then
      
    
      If R$ < Format(Now, "mm/dd/yyyy") Or Val(Right(R$, 4)) < Val(Format(Now, "yyyy")) Then
        MsgBox "Cannot select a date prior to today", 16, "Attention"
        lblfecha_cita1.Caption = ""
        GoTo salida
      End If
      
    Else
    
    
    
      If (ano1$ < ano2$ Or Val(Right(R$, 4)) < Val(Right(fecha_guardada$, 4))) And lblquote1.Caption <> "" Then
          MsgBox "Cannot select a date prior to saved date", 16, "Attention"
          GoTo salida
      ElseIf (ano1$ < Format(Now, "mm/dd/yyyy") Or Val(Right(R$, 4)) < Val(Format(Now, "yyyy"))) And lblquote1.Caption = "" Then
          MsgBox "Cannot select a date prior to today", 16, "Attention"
          GoTo salida
          
      ElseIf (ano1$ < lblappointment2.Caption Or Val(Right(R$, 4)) < Val(Right(lblappointment2.Caption, 4))) And lblquote1.Caption <> "" And cita = 1 Then
          MsgBox "Cannot select a date prior to second date", 16, "Attention"
          GoTo salida
          
      ElseIf (ano1$ < lblappointment3.Caption Or Val(Right(R$, 4)) < Val(Right(lblappointment3.Caption, 4))) And lblquote1.Caption <> "" And cita = 2 Then
          MsgBox "Cannot select a date prior to third date", 16, "Attention"
          GoTo salida
          
      End If
    
    
    End If
   
    lblfecha_cita1.Caption = R$
   
   
   
   'lblappointment2
  
salida:
Calendar1.Visible = False

End Sub






Private Sub Calendar1_LostFocus()
On Error Resume Next
Calendar1.Visible = False

End Sub


Private Sub Calendar2_DateClicked(inputDate As Date)
On Error Resume Next
     
    
     
    R$ = Format(inputDate, "mm/dd/yyyy")
    
    
            
   txtdate1.Text = R$
   
   
   
   'lblappointment2
  
salida:
Calendar2.Visible = False
End Sub

Private Sub Calendar3_DateClicked(inputDate As Date)
On Error Resume Next
     
    
     
    R$ = Format(inputDate, "mm/dd/yyyy")
    
    
            
   txtdate2.Text = R$
   
   
   
   'lblappointment2
  
salida:
Calendar3.Visible = False
End Sub


Private Sub cbo_Oficina1_Click()
On Error Resume Next


G$ = UCase$(cbo_Oficina1.List(cbo_Oficina1.ListIndex))


existe = 0
For z = 0 To cbo_Oficina1.ListCount
  If UCase$(matriz_oficina$(z, 0)) = G$ Then
      X = z
      existe = 1
      Exit For
  End If
Next z

If existe = 1 Then
  lbloficina.Caption = matriz_oficina$(X, 1)
Else
  lbloficina.Caption = ""
  cbo_Oficina1.ListIndex = -1
End If


Calendar1.Visible = False

End Sub


Private Sub cbo_time1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub cbocarrier_Click()
On Error Resume Next
If valido1 = 99 Then Exit Sub
carrier = cbocarrier.ListIndex


Calendar1.Visible = False



End Sub


Private Sub cboestado1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub cbostatus_gral1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub cbostatus_gral2_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub chk_agente_Click()
On Error Resume Next
carga_registros

End Sub

Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2

' **********   quitar
' agente = "CCF02"
 lblagente.Caption = agente
 
 If administrador$ = "N" Then
   btnemployees.Enabled = False
   btnOffices.Enabled = False
   
 End If
 

' *****************

modificada = 0
lblfecha.Caption = Format(Now, "mm/dd/yyyy")

busqueda_activada = 0
fecha_busqueda = 0



If full_access1$ = "Y" Then
  btnstatus.Visible = True
  btnlimpia_status3.Visible = True
Else
  btnstatus.Visible = False
  btnlimpia_status3.Visible = False
End If




tipo_sort = 0

Dim lRet As Long
    lRet = GetWindowLong(Me.hWnd, GWL_STYLE)
    lRet = lRet And Not (WS_MAXIMIZEBOX)
    lRet = SetWindowLong(Me.hWnd, GWL_STYLE, lRet)
    DesactivarMenu Me
    
    

Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1024  '1280
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 940 ' 1024
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = 0.78 '(Xpixels / DesignX)
        ScaleFactorY = 0.78 ' (Ypixels / DesignY)
      Else
        'ScaleFactorX = (Xpixels / DesignX)
        'ScaleFactorY = (Ypixels / DesignY)
      
        ScaleFactorX = 1280 / DesignX
        ScaleFactorY = 1024 / DesignY
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        Height = 9000 'Me.Height ' Remember the current size
        Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0

' carga tiempos del cbo_time1
cbo_time1.Clear
'cbo_time1.AddItem "07:00 am"
'cbo_time1.AddItem "07:15 am"
'cbo_time1.AddItem "07:30 am"
'cbo_time1.AddItem "07:45 am"
'cbo_time1.AddItem "08:00 am"
'cbo_time1.AddItem "08:15 am"
'cbo_time1.AddItem "08:30 am"
'cbo_time1.AddItem "08:45 am"
cbo_time1.AddItem "09:00 am"
cbo_time1.AddItem "09:15 am"
cbo_time1.AddItem "09:30 am"
cbo_time1.AddItem "09:45 am"
cbo_time1.AddItem "10:00 am"
cbo_time1.AddItem "10:15 am"
cbo_time1.AddItem "10:30 am"
cbo_time1.AddItem "10:45 am"
cbo_time1.AddItem "11:00 am"
cbo_time1.AddItem "11:15 am"
cbo_time1.AddItem "11:30 am"
cbo_time1.AddItem "11:45 am"
cbo_time1.AddItem "12:00 pm"
cbo_time1.AddItem "12:15 pm"
cbo_time1.AddItem "12:30 pm"
cbo_time1.AddItem "12:45 pm"
cbo_time1.AddItem "01:00 pm"
cbo_time1.AddItem "01:15 pm"
cbo_time1.AddItem "01:30 pm"
cbo_time1.AddItem "01:45 pm"
cbo_time1.AddItem "02:00 pm"
cbo_time1.AddItem "02:15 pm"
cbo_time1.AddItem "02:30 pm"
cbo_time1.AddItem "02:45 pm"
cbo_time1.AddItem "03:00 pm"
cbo_time1.AddItem "03:15 pm"
cbo_time1.AddItem "03:30 pm"
cbo_time1.AddItem "03:45 pm"
cbo_time1.AddItem "04:00 pm"
cbo_time1.AddItem "04:15 pm"
cbo_time1.AddItem "04:30 pm"
cbo_time1.AddItem "04:45 pm"
cbo_time1.AddItem "05:00 pm"
cbo_time1.AddItem "05:15 pm"
cbo_time1.AddItem "05:30 pm"
cbo_time1.AddItem "05:45 pm"
cbo_time1.AddItem "06:00 pm"
cbo_time1.AddItem "06:15 pm"
cbo_time1.AddItem "06:30 pm"
cbo_time1.AddItem "06:45 pm"
cbo_time1.AddItem "07:00 pm"
cbo_time1.AddItem "07:15 pm"
cbo_time1.AddItem "07:30 pm"
'cbo_time1.AddItem "07:45 pm"
'cbo_time1.AddItem "08:00 pm"
'cbo_time1.AddItem "08:15 pm"
'cbo_time1.AddItem "08:30 pm"
'cbo_time1.AddItem "08:45 pm"


cbostatus_gral1.Clear
cbostatus_gral1.AddItem "Sold (SD)"
cbostatus_gral1.AddItem "In process (IN)"
cbostatus_gral1.AddItem "Not Sold (NS)"
cbostatus_gral1.AddItem "Existing Customer (EC)"
cbostatus_gral1.AddItem "Commercial Quote (CQ)"
cbostatus_gral1.AddItem "No show (NSH)"


cbostatus_gral2.Clear
cbostatus_gral2.AddItem "Sold (SD)"
cbostatus_gral2.AddItem "In process (IN)"
cbostatus_gral2.AddItem "Not Sold (NS)"
cbostatus_gral2.AddItem "Existing Customer (EC)"
cbostatus_gral2.AddItem "Commercial Quote (CQ)"
cbostatus_gral2.AddItem "No show (NSH)"



cboestado1.Clear
cboestado1.AddItem "CA"
cboestado1.AddItem "TX"
cboestado1.AddItem "AZ"
cboestado1.AddItem "UT"
cboestado1.AddItem "NV"
cboestado1.AddItem "IL"


cbocarrier.Clear

cbocarrier.AddItem "AT&T"
cbocarrier.AddItem "Boost Mobile"
cbocarrier.AddItem "Cellular One"
cbocarrier.AddItem "Cingular"
cbocarrier.AddItem "Cricket"
cbocarrier.AddItem "Metro PCS"
cbocarrier.AddItem "Nextel"
cbocarrier.AddItem "PCS One"
cbocarrier.AddItem "Sprint"
cbocarrier.AddItem "T-Mobile"
cbocarrier.AddItem "U.S. Cellular"
cbocarrier.AddItem "Verizon"


'If regional >= 0 Then
' chk_agente.Enabled = True
'Else
'  chk_agente.Enabled = False
'End If



Conecta_SQL

carga_oficinas

carga_registros


End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1
End Sub



Public Sub limpia_campos()
On Error Resume Next
Calendar1.Visible = False
Calendar2.Visible = False
Calendar3.Visible = False

cita_registrada = -1
carrier = -1
lblquote1.Caption = ""
lbldate1.Caption = Format(Now, "mm/dd/yyyy")
txtcliente1.Text = ""
txtAfi1.Text = ""
txttelefono1.Text = ""
txtdireccion1.Text = ""
txtciudad1.Text = ""
cboestado1.ListIndex = -1
txtcp1.Text = ""
lblfecha_cita1.Caption = Format(Now, "mm/dd/yyyy")
cbo_time1.ListIndex = -1
lblstatus1.Caption = ""
cita = 0
op_cita(0).Picture = imgcita1b.Picture
op_cita(1).Picture = imgcita2.Picture
op_cita(2).Picture = imgcita3.Picture

cbostatus_gral1.Text = ""

lblappointment1.Caption = ""
lbltime_appointment1.Caption = ""

lblappointment2.Caption = ""
lblappointment3.Caption = ""
lbltime_appointment2.Caption = ""
lbltime_appointment3.Caption = ""
txtrecibo1.Text = ""
txtvendor1.Text = ""
txthwks1.Text = ""
cbostatus_gral1.ListIndex = -1
txtcomentarios1.Text = ""
txtCSR1.Text = ""
lbloficina.Caption = ""
cbo_Oficina1.ListIndex = -1
id_cita = 0
txtcomentario2.Visible = False
txtcomentario2.Text = ""

op_cita(0).Enabled = True
op_cita(1).Enabled = True
op_cita(2).Enabled = True
modificada = 0
btnsave1.Enabled = True
'cbostatus_gral1.Enabled = True
cbostatus_gral2.Enabled = True
cbostatus_gral2.ListIndex = -1
happy_face.Visible = False

' habilita campos
txttelefono1.Enabled = True
txtdireccion1.Enabled = True
txtciudad1.Enabled = True
cboestado1.Enabled = True
txtcp1.Enabled = True
cbo_Oficina1.Enabled = True
txtcliente1.Enabled = True
txtAfi1.Enabled = True
txtCSR1.Enabled = True
txtcomentarios1.Enabled = True
txtcelular.Text = ""
cbocarrier.ListIndex = -1

txtcelular.Enabled = True
cbocarrier.Enabled = True
txtturbo_quote.Text = ""
txtturbo_quote.Enabled = True

cboestado1.ListIndex = 1

If Dir$("c:\callcenter\CSR") <> "" Then
 nf = FreeFile
 Open "c:\callcenter\CSR" For Input Shared As #nf
 Lock #nf
 Line Input #nf, R$
 Unlock #nf
 Close #nf
 txtCSR1.Text = R$
End If



End Sub

Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   

 contraseña_ini$ = "PAieMu2DLBA6uNj86rSnCDpP"    '"admin"
 user_ini$ = "callc"   '"sa"
 bd_ini$ = "callcenter"   ' "CallCenter"
 server_ini$ = "justautocallcenter.couaea5kjoa1.us-west-1.rds.amazonaws.com"
 
 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub

Public Sub carga_oficinas()
On Error Resume Next

  
    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT nombre, abreviatura, direccion, cp FROM oficina"
   

    cbo_Oficina1.Clear
    Erase matriz_oficina$
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    Rs.Open sSelect, base, adOpenStatic

    
    ' define cuantos registros tiene la tabla
    columnas = Rs.Fields.Count
    filas = Rs.RecordCount
'Llenar las filas

       'For i = 0 To columnas
       ' grid1.TextMatrix(0, i) = Rs.Fields(i).Name
       'Next i
       
      'Llenar las filas
       For j = 1 To filas 'comenzamos en 1 porque el encabezado no se vuelve a llenar
      
              
              If Rs.Fields(0).value <> "" Then
                 cbo_Oficina1.AddItem Rs.Fields(0).value + " " + Rs.Fields(1).value
              End If
              
              matriz_oficina$(j, 0) = Rs.Fields(0).value
              matriz_oficina$(j, 1) = Rs.Fields(1).value
              matriz_oficina$(j, 2) = Rs.Fields(2).value
              matriz_oficina$(j, 3) = Rs.Fields(3).value
              
              
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
    
       total_reg = cbo_Oficina1.ListCount ' - 1
       'For t = 0 To total_reg
       '    matriz_oficina$(t, 0) = Left(cbo_Oficina1.List(t), Len(cbo_Oficina1.List(t)) - 3)
       '    matriz_oficina$(t, 1) = Right(cbo_Oficina1.List(t), 3)
           
      ' Next t
    
       cbo_Oficina1.Clear
       For t = 0 To total_reg
         If matriz_oficina$(t, 0) <> "" Then cbo_Oficina1.AddItem matriz_oficina$(t, 0)
       Next t
        
        
    
    Rs.Close
    
    
    
   
    
End Sub





Private Sub grid1_Click()
On Error Resume Next
  
  If grid1.Row = 0 Then Exit Sub
  
  carga_oficinas

  btnsave1.Enabled = True
  'cbostatus_gral1.Enabled = True
  cbostatus_gral2.Enabled = True
  happy_face.Visible = False
  cbostatus_gral1.ListIndex = -1
  cbostatus_gral2.ListIndex = -1
  cbostatus_gral1.Text = ""
  cbocarrier.ListIndex = -1
  fila = grid1.Row
  
  
  
   ' ************************************************************
  ' carga el ID


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
   
  
  
  'grid1.Col = 0
  'id_oficina = Val(grid1.Text)
  
  grid1.Col = 1
  lblquote1.Caption = grid1.Text
  
  grid1.Col = 2
  lblagente.Caption = agente
  
  grid1.Col = 3
  lbloficina.Caption = grid1.Text
  
  existe = 0
  For t = 0 To 99
    If RTrim(matriz_oficina$(t, 1)) = RTrim(lbloficina.Caption) Then
      For Y = 0 To cbo_Oficina1.ListCount - 1
         If RTrim(cbo_Oficina1.List(Y)) = RTrim(matriz_oficina$(t, 0)) Then
              cbo_Oficina1.ListIndex = Y
              existe = 1
              Exit For
         End If
      Next Y
    End If
    If existe = 1 Then Exit For
  Next t
  
  
  
  grid1.Col = 4
  lbldate1.Caption = grid1.Text
  
  
  
  grid1.Col = 5
  lblfecha_cita1.Caption = grid1.Text
  
  
  grid1.Col = 6
  cbo_time1.Text = grid1.Text
   
  
  grid1.Col = 7
  statusx$ = grid1.Text
  For z = 0 To cbostatus_gral1.ListCount - 1
    If cbostatus_gral1.List(z) = statusx$ Then
       cbostatus_gral1.ListIndex = z
       Exit For
    End If
  Next z
  
  
  grid1.Col = 8
  txtcliente1.Text = grid1.Text
  
  grid1.Col = 9
  lblappointment1.Caption = grid1.Text
  
  grid1.Col = 10
  lbltime_appointment1.Caption = grid1.Text
  
  
  
  grid1.Col = 11
  lblappointment2.Caption = grid1.Text
  
  grid1.Col = 12
  lbltime_appointment2.Caption = grid1.Text
  
  grid1.Col = 13
  lblappointment3.Caption = grid1.Text
  
  grid1.Col = 14
  lbltime_appointment3.Caption = grid1.Text
  
  grid1.Col = 15
  txttelefono1.Text = grid1.Text
  
  
  grid1.Col = 16
  txtdireccion1.Text = grid1.Text
  
  grid1.Col = 17
  txtciudad1.Text = grid1.Text
  
  grid1.Col = 18
  cboestado1.Text = grid1.Text
  
  grid1.Col = 19
  txtcp1.Text = grid1.Text
  
  grid1.Col = 20
  txtrecibo1.Text = grid1.Text
  
  grid1.Col = 21
  txtvendor1.Text = grid1.Text
  
  grid1.Col = 22
  txthwks1.Text = grid1.Text
  
  'grid1.Col = 21
  'cbostatus_gral1.Text = grid1.Text
  
  grid1.Col = 23
  txtcomentarios1.Text = grid1.Text
  
  grid1.Col = 24
  txtcomentarios2.Text = grid1.Text
  
  
  grid1.Col = 25
  txtCSR1.Text = grid1.Text
  
  grid1.Col = 26
  If Val(grid1.Text) = 0 Then
    txtturbo_quote.Text = ""
  Else
     txtturbo_quote.Text = grid1.Text
  End If
  
  
  
  
 
  Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idcita From CITAS where quote='" + lblquote1.Caption + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_cita = Rs(0)
            
                         
    Rs.Close
    

 
   
  ' ************************************************************
  ' carga el campo de AFI


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT afi From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtAfi1.Text = Rs(0)
            
                         
    Rs.Close
    
    
    
     ' ************************************************************
  ' carga la cita


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT cita1 From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    c1 = Rs(0)
           
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT cita2 From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    c2 = Rs(0)
           
                         
    Rs.Close
    
    
     Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT cita3 From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    c3 = Rs(0)
           
                         
    Rs.Close
   
    
 
 
op_cita(0).Picture = imgcita1.Picture
op_cita(1).Picture = imgcita2.Picture
op_cita(2).Picture = imgcita3.Picture

op_cita(0).Enabled = True
op_cita(1).Enabled = True
op_cita(2).Enabled = True


If c1 = 1 Then
    op_cita(0).Picture = imgcita1b.Picture
    cita = 0
    cita_registrada = 0
    'op_cita(0).Enabled = False
    'op_cita(1).Picture = imgcita2b.Picture
    'cita = 1

ElseIf c2 = 1 Then
    op_cita(1).Picture = imgcita2b.Picture
    cita = 1
    cita_registrada = 1
    op_cita(0).Enabled = False

ElseIf c3 = 1 Then
    op_cita(0).Enabled = False
    op_cita(1).Enabled = False


    op_cita(2).Picture = imgcita3b.Picture
    cita = 2
    cita_registrada = 2
End If




 ' ************************************************************
  ' carga el campo status_gral


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT status_gral From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    R$ = Rs(0)
            
                         
    Rs.Close
    
    
    If R$ <> "" Then
      If Left(R$, 4) = "Sold" And full_access1$ = "N" Then
       btnsave1.Enabled = False
       cbostatus_gral1.Enabled = False
       cbostatus_gral2.Enabled = False
       happy_face.Visible = True
      End If
    End If
    
    
    
    ' ************************************************************
  ' carga el campo status temporal


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT status From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    R$ = Rs(0)
    cbostatus_gral2.Text = R$
                         
    Rs.Close
    
    
        
    
    ' ************************************************************
  ' carga el campo comentario2


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT comentario2 From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    R$ = Rs(0)
            
                         
    Rs.Close
    
    txtcomentario2.Text = RTrim(R$)
    
    
    txtcomentario2.Visible = True
    
    
       ' ************************************************************
  ' carga el campo celular


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT celular From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    R$ = Rs(0)
            
                         
    Rs.Close
    
    txtcelular.Text = RTrim(R$)
    
    
       ' ************************************************************
  ' carga el campo carrier


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT sms From citas where idcita='" + Format(id_cita, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    R$ = Rs(0)
            
                         
    Rs.Close
    
    carrier = Val(RTrim(R$))
    
    If carrier >= 0 Then
      valido1 = 99
      cbocarrier.ListIndex = carrier
      valido1 = -1
    End If
    
    
    
    
    
    
    inhabilita_campos
     
End Sub

Private Sub op_carrier_Click(Index As Integer)
End Sub

Private Sub op_cita_Click(Index As Integer)
On Error Resume Next

If lblquote1.Caption = "" Then Exit Sub

If cita_registrada = 0 And Index = 2 Then
   Exit Sub
End If






op_cita(0).Picture = imgcita1.Picture
op_cita(1).Picture = imgcita2.Picture
op_cita(2).Picture = imgcita3.Picture

If Index = 0 Then
    op_cita(0).Picture = imgcita1b.Picture
     
ElseIf Index = 1 Then
    op_cita(1).Picture = imgcita2b.Picture

ElseIf Index = 2 Then
    op_cita(2).Picture = imgcita3b.Picture

End If

cita = Index


End Sub


Private Sub op_search_Click(Index As Integer)
On Error Resume Next

busqueda_activada = Index + 1

If Index = 4 Then
   If cbo_Oficina1.ListIndex >= 0 Then
       txtbusca.Text = cbo_Oficina1.List(cbo_Oficina1.ListIndex)
   End If
End If

End Sub



Private Sub op_showdate_Click(Index As Integer)
On Error Resume Next
fecha_busqueda = Index + 1

If Index = 2 Then
  txtdate1.Enabled = True
  txtdate2.Enabled = True
Else
  txtdate1.Enabled = False
  txtdate2.Enabled = False
End If




End Sub

Private Sub op_sort_Click(Index As Integer)
On Error Resume Next
tipo_sort = Index
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
lblhora.Caption = Format(Now, "hh:mm am/pm")

If seg >= 60 Then
  seg = 0
  carga_registros

  fecha_hoy$ = Format(Now, "mm/dd/yyyy")
  If lblfecha.Caption <> fecha_hoy$ Then
    End
  End If
  

End If




End Sub



Public Sub carga_quotes()

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
seg2 = seg2 + 1
If seg2 >= 30 Then
  Calendar1.Visible = False
  Calendar2.Visible = False
  Calendar3.Visible = False
  seg2 = 0
  Timer2.Enabled = False
End If


End Sub

Private Sub txtAfi1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtcelular_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub

Private Sub txtcelular_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc("-") Or KeyAscii = Asc("+") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If

End Sub


Private Sub txtciudad1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtcliente1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtcomentario2_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtcomentarios1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtcp1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub

Private Sub txtcp1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub



Public Sub inhabilita_campos()
On Error Resume Next

If full_access1$ = "N" Then
 txttelefono1.Enabled = False
 txtdireccion1.Enabled = False
 txtciudad1.Enabled = False
 cboestado1.Enabled = False
 txtcp1.Enabled = False

'cbo_Oficina1.Enabled = False
 txtcliente1.Enabled = False


 If txtCSR1.Text <> "" Then
  txtCSR1.Enabled = False
  txtAfi1.Enabled = False
 End If

 txtcomentarios1.Enabled = False

'txtcelular.Enabled = False
'cbocarrier.Enabled = False
'txtturbo_quote.Enabled = False
End If



End Sub

Private Sub txtdate1_GotFocus()
On Error Resume Next
Calendar2.Visible = True
Calendar3.Visible = False

End Sub

Private Sub txtdate1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


Private Sub txtdate2_GotFocus()
On Error Resume Next
Calendar3.Visible = True
Calendar2.Visible = False
End Sub

Private Sub txtdate2_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


Private Sub txtdireccion1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txthwks1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txttelefono1_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub

Private Sub txttelefono1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc("-") Or KeyAscii = Asc("+") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If


End Sub


Private Sub txtturbo_quote_Click()
On Error Resume Next
Calendar1.Visible = False
End Sub


Private Sub txtturbo_quote_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


Private Sub txtturbo_quote_LostFocus()
On Error Resume Next
txtturbo_quote.Font.Name = "Time new roman"
txtturbo_quote.Font.Size = 9
End Sub


