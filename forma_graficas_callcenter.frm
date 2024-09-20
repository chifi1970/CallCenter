VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form forma_graficas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Appointment Scheduler"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18105
   Icon            =   "forma_graficas_callcenter.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   -400
      TabIndex        =   70
      Top             =   11040
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   120
      ScaleHeight     =   7755
      ScaleWidth      =   14115
      TabIndex        =   1
      Top             =   1800
      Width           =   14175
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shown by: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7560
         TabIndex        =   74
         Top             =   4320
         Width           =   1815
         Begin Project1.lvButtons_H btnok 
            Height          =   375
            Left            =   1440
            TabIndex        =   78
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Caption         =   "OK"
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
            cBack           =   -2147483633
         End
         Begin VB.TextBox txtagente 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton op_agente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Only"
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
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton op_agente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ALL Agents"
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
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.PictureBox mensaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   5040
         ScaleHeight     =   1065
         ScaleWidth      =   3825
         TabIndex        =   42
         Top             =   2400
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Image Image6 
            Height          =   975
            Left            =   -120
            Picture         =   "forma_graficas_callcenter.frx":3336E
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label6 
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
            TabIndex        =   44
            Top             =   555
            Width           =   4095
         End
         Begin VB.Label Label5 
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
            TabIndex        =   43
            Top             =   120
            Width           =   3975
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   7440
         ScaleHeight     =   3255
         ScaleWidth      =   1935
         TabIndex        =   72
         Top             =   1560
         Width           =   1935
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   1635
         TabIndex        =   71
         Top             =   120
         Width           =   1695
         Begin VB.Image Image1 
            Height          =   1575
            Left            =   -120
            Picture         =   "forma_graficas_callcenter.frx":33D87
            Stretch         =   -1  'True
            Top             =   -240
            Width           =   2535
         End
      End
      Begin VB.PictureBox titulo_oficina 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5400
         ScaleHeight     =   585
         ScaleWidth      =   2745
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
         Begin VB.Label lbllugar 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Call Center"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.PictureBox by_region 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   5880
         ScaleHeight     =   4215
         ScaleWidth      =   1575
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   64
            Top             =   3360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "All offices"
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
            Mode            =   2
            Value           =   -1  'True
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   60
            Top             =   1920
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   61
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   62
            Top             =   2640
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   63
            Top             =   3000
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   65
            Top             =   3720
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "oficina"
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
            Mode            =   2
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "By Region"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   100
            Width           =   975
         End
      End
      Begin VB.OptionButton op_tipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5280
         Width           =   615
      End
      Begin VB.OptionButton op_tipo 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   2
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5280
         Width           =   615
      End
      Begin VB.OptionButton op_tipo 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   1
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4920
         Width           =   615
      End
      Begin VB.OptionButton op_tipo 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   0
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4920
         Value           =   -1  'True
         Width           =   615
      End
      Begin Project1.lvButtons_H btnexcel1 
         Height          =   855
         Left            =   12240
         TabIndex        =   56
         Top             =   3840
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
         Image           =   "forma_graficas_callcenter.frx":35D56
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnprint1 
         Height          =   735
         Left            =   11400
         TabIndex        =   55
         Top             =   3840
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
         Image           =   "forma_graficas_callcenter.frx":37B66
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   8640
         TabIndex        =   66
         Top             =   5040
         Width           =   4335
         Begin VB.ComboBox cbodia 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   320
            Width           =   615
         End
         Begin VB.OptionButton op_dia 
            BackColor       =   &H00FFFFFF&
            Caption         =   "One day"
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
            Index           =   1
            Left            =   960
            TabIndex        =   68
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton op_dia 
            BackColor       =   &H00FFFFFF&
            Caption         =   "The whole month"
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
            Index           =   0
            Left            =   960
            TabIndex        =   67
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.PictureBox accesototalx 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   45
         Top             =   -40
         Visible         =   0   'False
         Width           =   2535
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   46
            Text            =   "Combo1"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Region:"
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
            Left            =   600
            TabIndex        =   47
            Top             =   120
            Width           =   975
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   540
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboimpre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10560
         TabIndex        =   38
         Text            =   "Combo1"
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ComboBox cbomes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   4560
         Width           =   1095
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   5040
         ScaleHeight     =   1575
         ScaleWidth      =   735
         TabIndex        =   30
         Top             =   5040
         Width           =   735
         Begin VB.Image Image5 
            Height          =   735
            Left            =   0
            Picture         =   "forma_graficas_callcenter.frx":385BD
            Stretch         =   -1  'True
            Top             =   720
            Width           =   615
         End
         Begin VB.Image Image3 
            Height          =   615
            Left            =   0
            Picture         =   "forma_graficas_callcenter.frx":3A19F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3600
         ScaleHeight     =   615
         ScaleWidth      =   495
         TabIndex        =   29
         Top             =   720
         Width           =   495
         Begin VB.Image Image2 
            Height          =   495
            Left            =   120
            Picture         =   "forma_graficas_callcenter.frx":3BE90
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
      End
      Begin MSChart20Lib.MSChart grafica3 
         Height          =   3375
         Left            =   0
         OleObjectBlob   =   "forma_graficas_callcenter.frx":3C488
         TabIndex        =   28
         Top             =   4320
         Width           =   5175
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   9360
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   17
         Top             =   840
         Width           =   2055
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status Gral. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   11640
         ScaleHeight     =   1935
         ScaleWidth      =   1215
         TabIndex        =   9
         Top             =   960
         Width           =   1215
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   25
            Top             =   1600
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   24
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   23
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   22
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   21
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lbl_cant_status 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   19
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Not rated"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   360
            TabIndex        =   16
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NSH"
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
            Index           =   5
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CQ"
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
            Index           =   4
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "EC"
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
            Index           =   3
            Left            =   360
            TabIndex        =   13
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NS"
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
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "IN"
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
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   375
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   6
            Left            =   120
            Top             =   1680
            Width           =   135
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FFFF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   5
            Left            =   120
            Top             =   1360
            Width           =   135
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   4
            Left            =   120
            Top             =   1120
            Width           =   135
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   3
            Left            =   120
            Top             =   880
            Width           =   135
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   120
            Top             =   640
            Width           =   135
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   400
            Width           =   135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SD"
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
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   375
         End
         Begin VB.Shape Shape3 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   160
            Width           =   135
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year"
         Height          =   615
         Left            =   10320
         TabIndex        =   5
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton op_year 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2019"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton op_year 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2018"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin MSChart20Lib.MSChart grafica1 
         Height          =   3975
         Left            =   0
         OleObjectBlob   =   "forma_graficas_callcenter.frx":3E420
         TabIndex        =   4
         Top             =   600
         Width           =   5775
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   1815
         Left            =   6000
         TabIndex        =   2
         Top             =   5760
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   16777215
         BorderStyle     =   0
         Appearance      =   0
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
      Begin MSChart20Lib.MSChart grafica2 
         Height          =   2895
         Left            =   7080
         OleObjectBlob   =   "forma_graficas_callcenter.frx":3FE2B
         TabIndex        =   8
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order by:"
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
         Left            =   11400
         TabIndex        =   73
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Printers:"
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
         Left            =   9960
         TabIndex        =   39
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Order by:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   11640
         TabIndex        =   34
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Month:"
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
         Left            =   9600
         TabIndex        =   32
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   200
         Width           =   8295
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11520
      Top             =   9720
   End
   Begin Project1.lvButtons_H btncallcenter 
      Height          =   1335
      Left            =   240
      TabIndex        =   48
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
      ImgAlign        =   4
      Image           =   "forma_graficas_callcenter.frx":42F2D
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnvendor 
      Height          =   1335
      Left            =   1560
      TabIndex        =   49
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
      Image           =   "forma_graficas_callcenter.frx":449A4
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnReports 
      Height          =   1335
      Left            =   2880
      TabIndex        =   50
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
      Value           =   -1  'True
      Image           =   "forma_graficas_callcenter.frx":47E7C
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnemployees 
      Height          =   1335
      Left            =   4200
      TabIndex        =   51
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
      Image           =   "forma_graficas_callcenter.frx":4A418
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnOffices 
      Height          =   1335
      Left            =   5520
      TabIndex        =   52
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
      Image           =   "forma_graficas_callcenter.frx":4C6A9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnexit 
      Height          =   1335
      Left            =   12840
      TabIndex        =   53
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
      Image           =   "forma_graficas_callcenter.frx":4E931
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Image Image7 
      Height          =   12855
      Left            =   14520
      Picture         =   "forma_graficas_callcenter.frx":50B20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1240
      TabIndex        =   0
      Top             =   11880
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   -840
      Top             =   -120
      Width           =   18015
   End
End
Attribute VB_Name = "forma_graficas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer
Dim id_employee As Integer, seg As Integer, oficina$, activo$, CallCenter$, manager$, tipo_manager As Integer, region As Integer, admon$, year1 As Integer, region_parcial As Integer
Dim mes_elegido As Integer, tipo_orden$, mes_o_dia As Integer, tipo_agente As Integer
Dim xprint As Printer

Dim matriz_oficina$(100, 4)
Dim matriz$(100, 2)


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)




Public Sub carga_impresoras()
On Error Resume Next

Dim cImprGen As String
    cImprGen = cboimpre.Text
    
cboimpre.Clear
    
    
If Dir$(ruta$ + "printer") <> "" Then
 nf = FreeFile
 Open ruta$ + "printer" For Input Shared As #nf
 Lock #nf
 Line Input #nf, P1$
 Line Input #nf, P2$
 Unlock #nf
 Close #nf
 
 cImprGen = P1$
 cboimpre.Text = P1$

End If
    
    
    
    
For Each xprint In Printers
           If xprint.DeviceName = cImprGen Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next
        
        
        
For Each xprint In Printers
        cboimpre.AddItem xprint.DeviceName
Next
        
        
nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
 For t = 0 To cboimpre.ListCount - 1
   If cboimpre.List(t) = Printer.DeviceName Then
       cboimpre.ListIndex = t
       Exit For
   End If
 Next t
        
        
        
        
End Sub

Public Sub enca()
On Error Resume Next

Printer.Print Space(1)

Printer.Font.Name = "Courier new"
Printer.FontSize = 14
Printer.Print Space(2) + "   Appointments                                       " + cbomes.List(cbomes.ListIndex) + "/" + Format(year1, "0000")
Printer.FontSize = 8

  Printer.Print Space(2) + "------------------------------------------------------------------------------------------------------------------------"
  Printer.Print Space(2) + " Quote  Office             Date        Status                Customer               Phone           CSR   Quote(TR)"
  Printer.Print Space(2) + "------------------------------------------------------------------------------------------------------------------------"


Printer.Print Space(1)

End Sub
Public Sub carga_registros()
On Error Resume Next


  
    ' Para la cadena de selección
    Dim sSelect As String
    
    fecha_hoy$ = Format(Now, "mm/dd/yyyy")
    
    
    ordena_por$ = "quote"
    
    
    
    If busqueda_activada = 0 Then
    
        Select Case fecha_busqueda
        Case 0
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_hoy$ + "' or fecha_cita2='" + fecha_hoy$ + "' or fecha_cita3='" + fecha_hoy$ + "' order by " + ordena_por$ + " asc"
        Case 1
              dia_ayer = Val(Format(Now, "y"))
              fecha_ayer$ = Format(dia_ayer, "mm/dd")
              fecha_ayer$ = fecha_ayer$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_ayer$ + "' or fecha_cita2='" + fecha_ayer$ + "' or fecha_cita3='" + fecha_ayer$ + "' order by " + ordena_por$ + " asc"
        Case 2
        dia_manana = Val(Format(Now, "y")) + 2
              fecha_manana$ = Format(dia_manana, "mm/dd")
              fecha_manana$ = fecha_manana$ + "/" + Format(Now, "yyyy")
              
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1='" + fecha_manana$ + "' or fecha_cita2='" + fecha_manana$ + "' or fecha_cita3='" + fecha_manana$ + "' order by " + ordena_por$ + " asc"
        
        Case 3
                            
              f1$ = Format(txtdate1.Text, "mm/dd/yyyy")
              f2$ = Format(txtdate2.Text, "mm/dd/yyyy")
              
              f1$ = "convert(datetime, '" + f1$ + "')"
              f2$ = "convert(datetime, '" + f2$ + "')"
              
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
              sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE fecha_cita1 between " + f1$ + " and " + f2$ + " or fecha_cita2 between " + f1$ + " and " + f2$ + " or fecha_cita3 between " + f1$ + " and " + f2$ + " order by " + ordena_por$ + " asc"
        
        End Select
   
    Else
     
         Select Case busqueda_activada
         Case 1
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE cliente like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 2
           
           
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE telefono like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         Case 3
           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE quote='" + UCase(txtbusca.Text) + "' order by " + ordena_por$ + " asc"
           
         Case 4
             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE direccion like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"
         
         End Select
         
         
         
         
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
                 grid1.TextMatrix(j, i) = Rs.Fields(i).value
              End If
              
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
    
    
    
    
    
    'idcita,  estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE cliente like '%" + UCase(txtbusca.Text) + "%' order by " + ordena_por$ + " asc"

    
    ' asigna anchos de columnas
    grid1.ColWidth(0) = 900   ' idcita
    grid1.ColWidth(1) = 1000  ' quote
    grid1.ColWidth(2) = 750  ' agente
    grid1.ColWidth(3) = 750  ' oficina
    grid1.ColWidth(4) = 1750  'fecha
    grid1.ColWidth(5) = 1450 ' fecha appointment
    grid1.ColWidth(6) = 820  '  hora cita
    grid1.ColWidth(7) = 1800  ' status
    grid1.ColWidth(8) = 3000  ' customer
    grid1.ColWidth(9) = 1450 ' appointment 2
    grid1.ColWidth(10) = 820   ' hora 2
    grid1.ColWidth(11) = 1450 ' appointment 3
    grid1.ColWidth(12) = 820   ' hora 3
    grid1.ColWidth(13) = 1400   ' telefono
    grid1.ColWidth(14) = 3200  ' direccion
    grid1.ColWidth(15) = 2000  ' ciudad
    grid1.ColWidth(16) = 600  ' estado
    grid1.ColWidth(17) = 700  ' zip
    grid1.ColWidth(18) = 1200  ' recibo
    grid1.ColWidth(19) = 1200  '  vendor
    grid1.ColWidth(20) = 1200  '  hwks
   
    grid1.ColWidth(21) = 3000  ' comen1
    grid1.ColWidth(22) = 3000  ' comen 2
    grid1.ColWidth(23) = 800  '  csr
    grid1.ColWidth(24) = 1000  ' quote turborater
    
    
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
    grid1.Text = "Appointment 1"
    
    grid1.Col = 6
    grid1.Text = "Time 1"
    
    grid1.Col = 7
    grid1.Text = "Status"
    
    grid1.Col = 8
    grid1.Text = "Customer"
    
    grid1.Col = 9
    grid1.Text = "Appointment 2"
    
    grid1.Col = 10
    grid1.Text = "Time 2"
    
    grid1.Col = 11
    grid1.Text = "Appointment 3"
    
    grid1.Col = 12
    grid1.Text = "Time 3"
    
    grid1.Col = 13
    grid1.Text = "Phone"
    
    grid1.Col = 14
    grid1.Text = "Address"
    
    grid1.Col = 15
    grid1.Text = "City"
    
    grid1.Col = 16
    grid1.Text = "State"
    
    grid1.Col = 17
    grid1.Text = "ZIP"
    
    grid1.Col = 18
    grid1.Text = "Receipt"
    
    grid1.Col = 19
    grid1.Text = "Vendor"
    
    grid1.Col = 20
    grid1.Text = "Hwks"
    
    'grid1.Col = 21
    'grid1.Text = "Status_gral."
    
    grid1.Col = 21
    grid1.Text = "Comments"
    
    grid1.Col = 22
    grid1.Text = "Comments"
        
    grid1.Col = 23
    grid1.Text = "CSR"
    
    grid1.Col = 24
    grid1.Text = "Quote TR"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub

Public Sub carga_oficinasxregion()

On Error Resume Next


  
    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT idoficina, abreviatura, nombre FROM oficina where region='" + STR(region_zona) + "' order by abreviatura asc"
   


    
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
       
       contador = 0
      'Llenar las filas
       For j = 1 To filas 'comenzamos en 1 porque el encabezado no se vuelve a llenar
          contador = contador + 1
          grid1.TextMatrix(j, 0) = Format(contador, "####0")
          For i = 1 To columnas
              grid1.TextMatrix(j, i) = Rs.Fields(i).value
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
    
    
    
    ' asigna anchos de columnas
    grid1.ColWidth(0) = 900
    grid1.ColWidth(1) = 900
    grid1.ColWidth(2) = 900
    
    ' cambia los titulos del GRID
    grid1.Row = 0
    
    grid1.Col = 0
    grid1.Text = ""
    grid1.RowHeight(0) = 600
    grid1.ColAlignment(0) = 4   ' 1=izq   4=centro  7=derecha
    grid1.ColAlignment(1) = 1
    grid1.ColAlignment(2) = 1
      
    
    Erase matriz$
    c = 0
    
    For t = 1 To grid1.rows - 1
      grid1.Col = 1
      grid1.Row = t
      abrev$ = grid1.Text
      grid1.Col = 2
      n$ = grid1.Text
      c = c + 1
      matriz$(c, 0) = n$
      matriz$(c, 1) = abrev$
      
      op_oficina(t).Caption = n$
      op_oficina(t).Visible = True
    Next t
   
    
    
    
    
    
    
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
      
              
              If Rs.Fields(0).value <> "" Then cbooficinas.AddItem Rs.Fields(0).value + " " + Rs.Fields(1).value
              
              matriz_oficina$(j, 0) = Rs.Fields(0).value
              matriz_oficina$(j, 1) = Rs.Fields(1).value
              matriz_oficina$(j, 2) = Rs.Fields(2).value
              matriz_oficina$(j, 3) = Rs.Fields(3).value
              
              
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
    
       total_reg = cbooficinas.ListCount - 1
       'For t = 0 To total_reg
       '    matriz_oficina$(t, 0) = Left(cbo_Oficina1.List(t), Len(cbo_Oficina1.List(t)) - 3)
       '    matriz_oficina$(t, 1) = Right(cbo_Oficina1.List(t), 3)
           
      ' Next t
    
       cbooficinas.Clear
       For t = 0 To total_reg
         If matriz_oficina$(t, 0) <> "" Then cbooficinas.AddItem matriz_oficina$(t, 0)
       Next t
        
        
    
    Rs.Close
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




Private Sub btncallcenter_Click()
On Error Resume Next
base.Close

Load forma_main
forma_main.Show
Unload Me

End Sub



Private Sub btnemployees_Click()
On Error Resume Next
base.Close

Load forma_employees
forma_employees.Show
Unload Me
End Sub

Private Sub btnexcel1_Click()
On Error Resume Next
Dim sData As String
 
If grid1.rows = 1 Then Exit Sub

' carga info en grid1


archivo$ = "c:\callcenter\Report1.xlsx"
Kill archivo$

If Dir$(archivo$) <> "" Then
  MsgBox "Please, close the file " + archivo$ + " and try it again.", 64, "Attention"
  Exit Sub
End If


mensaje.Visible = True
mensaje.Refresh


sData = "Quote" & vbTab & "Office" & vbTab & "date" & vbTab & "Time" & vbTab & "Status" & vbTab & "Customer" & vbTab & "Phone" & vbTab & "Address" & vbTab & "City" _
& vbTab & "State" & vbTab & "Zip" & vbTab & "CSR" & vbTab & "Quote TR" & vbCr


  
 
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
  
  For Y = 1 To 13
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
     'MsgBox "You have open the file " + archivo$ + ". It couldn't save anything. Close it and try again.", 16, "Attention"
     'oExcel.Quit
     'mensaje.Visible = False

     'Exit Sub
     
   End If
   
   oExcel.Quit
   
   mensaje.Visible = False
   
   If Dir$("C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE") <> "" Then
      R$ = Shell("C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE c:\callcenter\Report1.xlsx", vbNormalFocus)
   Else
      R$ = Shell("C:\Program Files\Microsoft Office\Office15\EXCEL.EXE c:\callcenter\Report1.xlsx", vbNormalFocus)
   End If
   
   
   
   MsgBox "The file named " + archivo$ + " was created successfully", 64, "Attention"
   

End Sub

Private Sub btnexit_Click()
On Error Resume Next
base.Close
End
End Sub

















Private Sub btnOffices_Click()
On Error Resume Next
base.Close

Load forma_oficinas
forma_oficinas.Show
Unload Me
End Sub



Private Sub btnok_Click()
calcula_datos_mes

End Sub

Private Sub btnprint1_Click()
On Error Resume Next
If grid1.rows = 1 Then Exit Sub

R$ = MsgBox("Do you wish to print this report?", 4, "Attention")
If R$ = "7" Then Exit Sub


reporte = 1
enca



c = 0

paginas = (grid1.rows - 1) / 50
pag1 = Int(paginas)
residuo = paginas - pag1
If paginas < 1 Then paginas = 0
If residuo > 0 Then paginas = paginas + 1
pag1 = 1


For t = 1 To grid1.rows - 1
  grid1.Row = t
  grid1.Col = 0
  num = Val(grid1.Text)
  
  '  sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE oficina='" + oficina_autorizada$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + " or fecha_cita2 between " + f1$ + " and " + f2$ + " or fecha_cita3 between " + f1$ + " and " + f2$ + ") order by quote asc"

  
  grid1.Col = 1
  quotex$ = grid1.Text
  
  grid1.Col = 2
  oficinax$ = grid1.Text
  
  grid1.Col = 3
  fechax$ = grid1.Text
  
  grid1.Col = 4
  horax$ = grid1.Text
  
  grid1.Col = 5
  statusx$ = grid1.Text
  If statusx$ = "" Then statusx$ = "     "
  
  
  grid1.Col = 6
  clientex$ = grid1.Text
  
  grid1.Col = 7
  telefonox$ = grid1.Text
  tel1$ = Right(telefonox$, 10)
  telefonox$ = "(" + Left(tel1$, 3) + ")" + Mid$(tel1$, 4, 3) + "-" + Right(tel1$, 4)
  
  grid1.Col = 8
  direccionx$ = grid1.Text
  
  grid1.Col = 9
  ciudadx$ = grid1.Text
  
  grid1.Col = 10
  estadox$ = grid1.Text
  
  grid1.Col = 11
  cpx$ = grid1.Text
  
  grid1.Col = 12
  csrx$ = grid1.Text
  If csrx$ = "" Then csrx$ = "   "
  
  
  grid1.Col = 13
  trx$ = grid1.Text
  If trx$ = "0" Then trx$ = " "
  
  
  
  
  Printer.Print Space(3) + Format(quotex$, "!@@@@@@") + Space(1) + Format(Left(oficinax$, 18), "!@@@@@@@@@@@@@@@@@@") + Space(1) + fechax$ + Space(2) + Format(statusx$, "!@@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(Left(clientex$, 22), "!@@@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(telefonox$, "!@@@@@@@@@@@@@@@") + Space(2) + Format(csrx$, "!@@@@") + Space(2) + Format(trx$, "!@@@@@@@@@@@@@@@")
  c = c + 1
  
  If c = 50 Then
    c = 0
    Printer.Print Space(1)
    Printer.Print Space(5) + "Page " + Format(pag1, "###0") + "/" + Format(paginas, "###0")
    pag1 = pag1 + 1
    Printer.NewPage
    enca
    
  End If
  
Next t



If pag1 < paginas Then
  For Y = c To 50
    Printer.Print Space(1)
  Next Y
  pagi1 = pag1 + 1
  Printer.Print Space(5) + "Page " + Format(pag1, "###0") + "/" + Format(paginas, "###0")
End If


Printer.EndDoc

MsgBox "The report was printed", 64, "Attention"


End Sub

Private Sub btnvendor_Click()
On Error Resume Next
base.Close

Load forma_vendor
forma_vendor.Show
Unload Me
End Sub



Private Sub cbodia_Click()
On Error Resume Next


calcula_datos_mes

End Sub


Private Sub cboimpre_Click()
On Error Resume Next


For Each xprint In Printers
           If xprint.DeviceName = cboimpre.Text Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next


nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
End Sub


Private Sub cbomes_Click()
On Error Resume Next
mes_elegido = cbomes.ListIndex + 1

cbodia.Clear
Select Case mes_elegido
Case 1, 3, 5, 7, 8, 10, 12
  dias = 31
Case 2
  R = year1 / 4
  res = R - Int(R)
  If res = 0 Then
    dias = 29
  Else
    dias = 28
  End If
  
Case 4, 6, 9, 11
  dias = 30
End Select

For t = 1 To dias
  cbodia.AddItem t
Next t

op_dia(0).value = True

calcula_datos_mes
End Sub


Private Sub Combo1_Click()
On Error Resume Next
For Y = 1 To 9
  op_oficina(Y).Visible = False
Next Y

op_oficina(0).value = True


region_zona = Combo1.ListIndex + 1
carga_oficinasxregion





region_parcial = 0
calcula_citasXmes

If cbomes.ListIndex >= 0 Then calcula_datos_mes

End Sub


Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2

lblfecha.Caption = Format(Now, "mm/dd/yyyy")


'If gerente$ = "N" Then
'  grafica1.TitleText = "Appointments by Agent"
'Else
'  grafica1.TitleText = "Appointments by Office"
'End If






year1 = Format(Now, "yyyy")
op_year(0).Caption = Format(year1 - 1, "0000")
op_year(1).Caption = Format(year1, "0000")

If callcenter1$ = "Y" Then
  btncallcenter.Enabled = True
  btnvendor.Enabled = False
  
  lbllugar.Caption = "Call Center"
Else
  btncallcenter.Enabled = False
  btnvendor.Enabled = True
 
  lbllugar.Caption = oficina_autorizada$
End If


If administrador$ = "Y" Then
 btnemployees.Enabled = True
 btnOffices.Enabled = True
Else
 btnemployees.Enabled = False
 btnOffices.Enabled = False
End If



region_parcial = 0
carga_impresoras

If regional = 1 Then
  by_region.Visible = True
  titulo_oficina.Visible = False
Else
  by_region.Visible = False
  titulo_oficina.Visible = True
  
  
End If


If full_access1$ = "Y" Then
  accesototalx.Visible = True
Else
  accesototalx.Visible = False
End If



tipo_orden$ = "order by quote asc"
mes_elegido = 0

Dim lRet As Long
    lRet = GetWindowLong(Me.hWnd, GWL_STYLE)
    lRet = lRet And Not (WS_MAXIMIZEBOX)
    lRet = SetWindowLong(Me.hWnd, GWL_STYLE, lRet)
    DesactivarMenu Me
    
 id_employee = 0

Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
       'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1024 '1280 '1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 940 ' 1024 '800
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

Conecta_SQL

cbomes.Clear
cbomes.AddItem "January"
cbomes.AddItem "February"
cbomes.AddItem "March"
cbomes.AddItem "April"
cbomes.AddItem "May"
cbomes.AddItem "June"
cbomes.AddItem "July"
cbomes.AddItem "August"
cbomes.AddItem "September"
cbomes.AddItem "October"
cbomes.AddItem "November"
cbomes.AddItem "December"


carga_oficinas

carga_oficinasxregion
' carga_registros

calcula_citasXmes

Combo1.Clear
For t = 1 To 30
  Combo1.AddItem t

Next t

Combo1.ListIndex = (region_zona - 1)


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

End Sub




Private Sub grafica1_GotFocus()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub grafica2_GotFocus()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub grafica3_GotFocus()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub grid1_Click()
On Error Resume Next
  carga_oficinas

  btnsave1.Enabled = True
  cbostatus_gral1.Enabled = True
  happy_face.Visible = False
  cbostatus_gral1.ListIndex = -1
    cbocarrier.ListIndex = -1

  fila = grid1.Row
  
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
  lblappointment2.Caption = grid1.Text
  
  grid1.Col = 10
  lbltime_appointment2.Caption = grid1.Text
  
  grid1.Col = 11
  lblappointment3.Caption = grid1.Text
  
  grid1.Col = 12
  lbltime_appointment3.Caption = grid1.Text
  
  grid1.Col = 13
  txttelefono1.Text = grid1.Text
  
  
  grid1.Col = 14
  txtdireccion1.Text = grid1.Text
  
  grid1.Col = 15
  txtciudad1.Text = grid1.Text
  
  grid1.Col = 16
  cboestado1.Text = grid1.Text
  
  grid1.Col = 17
  txtcp1.Text = grid1.Text
  
  grid1.Col = 18
  txtrecibo1.Text = grid1.Text
  
  grid1.Col = 19
  txtvendor1.Text = grid1.Text
  
  grid1.Col = 20
  txthwks1.Text = grid1.Text
  
  'grid1.Col = 21
  'cbostatus_gral1.Text = grid1.Text
  
  grid1.Col = 21
  txtcomentarios1.Text = grid1.Text
  
  grid1.Col = 22
  txtcomentarios2.Text = grid1.Text
  
  grid1.Col = 23
  txtCSR1.Text = grid1.Text
  
  grid1.Col = 24
  If Val(grid1.Text) = 0 Then
    txtturbo_quote.Text = ""
  Else
     txtturbo_quote.Text = grid1.Text
  End If
  
  
  
  
  ' ************************************************************
  ' carga el ID


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
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
      If Left(R$, 4) = "Sold" Then
       btnsave1.Enabled = False
       cbostatus_gral1.Enabled = False
       happy_face.Visible = True
      End If
    End If
    
    
    
    
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
    
    
    
    
    
    
    
     

  
End Sub











Private Sub op_agente_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
  btnok.Enabled = False
  txtagente.Enabled = False
Else
  btnok.Enabled = True
  txtagente.Enabled = True
End If

tipo_agente = Index
calcula_datos_mes


End Sub

Private Sub op_dia_Click(Index As Integer)
On Error Resume Next

mes_o_dia = Index

If Index = 0 Then
  cbodia.ListIndex = -1
  cbodia.Enabled = False
Else
  cbodia.Enabled = True
End If

calcula_datos_mes


End Sub

Private Sub op_oficina_Click(Index As Integer)
On Error Resume Next

region_parcial = Index
calcula_citasXmes

If cbomes.ListIndex >= 0 Then calcula_datos_mes

End Sub

Private Sub op_tipo_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
  tipo_orden$ = "order by quote asc"
Case 1
  tipo_orden$ = "order by fecha_cita1 asc"
Case 2
  tipo_orden$ = "order by oficina asc"
Case 3
  tipo_orden$ = "order by status_gral asc"
End Select

calcula_datos_mes
      

End Sub

Private Sub op_year_Click(Index As Integer)
On Error Resume Next
year1 = Val(op_year(Index).Caption)
calcula_citasXmes
End Sub

Private Sub Timer1_Timer()
seg = seg + 1
lblhora.Caption = Format(Now, "hh:mm am/pm")
End Sub











Public Sub calcula_citasXmes()
On Error Resume Next

grafica1.RowCount = 1


Dim data_status(7) As Integer, data_sold(12, 2) As Integer
Erase data_status
Erase data_sold

'For t = 0 To 11
' grafica1.Column = t + 1
' grafica1.Data = t + 1
 
'Next t



' inicializa graficas
For t = 1 To 12
   grafica1.Column = t
   grafica1.Data = 0
   grafica3.Column = t
   grafica3.Data = 0
Next t

For z = 1 To 7
 grafica2.Column = z
 grafica2.Data = 0
 lbl_cant_status(z - 1).Caption = ""
Next z
  
grafica1.Refresh
grafica2.Refresh
grafica3.Refresh



For t = 1 To 12

  Select Case t
  Case 1
     date1$ = "01/01/" + Format(year1, "0000")
     date2$ = "01/31/" + Format(year1, "0000")
  Case 2
     res = (year1 / 4) - Int(year1 / 4)
     If res = 0 Then
        date1$ = "02/01/" + Format(year1, "0000")
        date2$ = "02/29/" + Format(year1, "0000")
     Else
        date1$ = "02/01/" + Format(year1, "0000")
        date2$ = "02/28/" + Format(year1, "0000")
     End If
        
  
  Case 3
     date1$ = "03/01/" + Format(year1, "0000")
     date2$ = "03/31/" + Format(year1, "0000")
  Case 4
     date1$ = "04/01/" + Format(year1, "0000")
     date2$ = "04/30/" + Format(year1, "0000")
  Case 5
     date1$ = "05/01/" + Format(year1, "0000")
     date2$ = "05/31/" + Format(year1, "0000")
  Case 6
     date1$ = "06/01/" + Format(year1, "0000")
     date2$ = "06/30/" + Format(year1, "0000")
  Case 7
     date1$ = "07/01/" + Format(year1, "0000")
     date2$ = "07/31/" + Format(year1, "0000")
  Case 8
     date1$ = "08/01/" + Format(year1, "0000")
     date2$ = "08/31/" + Format(year1, "0000")
  Case 9
     date1$ = "09/01/" + Format(year1, "0000")
     date2$ = "09/30/" + Format(year1, "0000")
  Case 10
     date1$ = "10/01/" + Format(year1, "0000")
     date2$ = "10/31/" + Format(year1, "0000")
  Case 11
     date1$ = "11/01/" + Format(year1, "0000")
     date2$ = "11/30/" + Format(year1, "0000")
  Case 12
     date1$ = "12/01/" + Format(year1, "0000")
     date2$ = "12/31/" + Format(year1, "0000")
  End Select
     
     
     

  If regional <= 0 Then
                                   
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              
              If f1$ = "" Then
                Exit Sub
              End If
              
              If callcenter1$ = "N" Then
                 sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE oficina='" + oficina_autorizada$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
              Else
                 sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + agente + "' and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
              End If
    
    
  Else
   ' si es un regional entonces            region_zona
   ' ********************************************************************************************************
                            
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
             ' If callcenter1$ = "N" Then
                 
                 ' selecciona oficina individual
                 
                 If region_parcial = 0 Then
                           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE region='" + Format(region_zona, "#0") + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                 Else
                           oficina_parcial$ = matriz$(region_parcial, 0)
                           sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                    
                 
                 End If
                           
             ' Else
              
             '     If region_parcial = 0 Then
                      
             '             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE agente='" + agente + "' and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
             '     Else
             '             oficina_parcial$ = matriz$(region_parcial, 0)
             '             sSelect = "SELECT idcita, quote, agente, oficina,fecha,fecha_cita1, hora_cita1,status_gral,cliente,fecha_citax, hora_citax,fecha_cita2, hora_cita2, fecha_cita3,hora_cita3, telefono, direccion, ciudad, estado,cp, recibo, vendor, hwks,  comentario, comentario2, csr, quote_turborater FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                  
             '     End If
                  
             ' End If
        
   
 ' **************************************************************************************
    
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
                 grid1.TextMatrix(j, i) = Rs.Fields(i).value
              End If
              
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
 
 grafica1.Column = t
 grafica1.Data = grid1.rows - 1
 
 
 
 ' calcula status
 ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 
  If regional <= 0 Then
                                   
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              
              If f1$ = "" Then
                Exit Sub
              End If
              
              If callcenter1$ = "N" Then
                 sSelect = "SELECT idcita,status_gral FROM citas WHERE oficina='" + oficina_autorizada$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
              Else
                 sSelect = "SELECT idcita,status_gral FROM citas WHERE agente='" + agente + "' and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
              End If
    
    
  Else
   ' si es un regional entonces            region_zona
   ' ********************************************************************************************************
                            
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
              'If callcenter1$ = "N" Then
                 If region_parcial = 0 Then
                           sSelect = "SELECT idcita,status_gral FROM citas WHERE region='" + Format(region_zona, "#0") + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                 Else
                           oficina_parcial$ = matriz$(region_parcial, 0)
                           sSelect = "SELECT idcita,status_gral FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                 
                 End If
              
              
              'Else
              '   If region_parcial = 0 Then
               '           sSelect = "SELECT idcita,status_gral FROM citas WHERE agente='" + agente + "' and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
              '   Else
              '             oficina_parcial$ = matriz$(region_parcial, 0)
              '             sSelect = "SELECT idcita,status_gral FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") order by quote asc"
                 
              '   End If
                  
              'End If
        
   
 ' **************************************************************************************
    
  End If
 
 
 
 
    
    ' El recordset para acceder a los datos
    
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
                 grid1.TextMatrix(j, i) = Rs.Fields(i).value
              End If
              
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
 
 
 
    grid1.Col = 1
    For z = 1 To (grid1.rows - 1)
      If grid1.rows = 1 Then Exit For
      'If grid1.Col = 0 Then Exit For
      grid1.Row = z
      
      Select Case grid1.Text
      Case "Sold (SD)"
          data_status(0) = data_status(0) + 1
          data_sold(t - 1, 0) = data_sold(t - 1, 0) + 1
      Case "In process (IN)"
          data_status(1) = data_status(1) + 1
      Case "Not Sold (NS)"
          data_status(2) = data_status(2) + 1
          data_sold(t - 1, 1) = data_sold(t - 1, 1) + 1
      Case "Existing Customer (EC)"
          data_status(3) = data_status(3) + 1
      Case "Commercial Quote (CQ)"
          data_status(4) = data_status(4) + 1
      Case "No show (NSH)"
          data_status(5) = data_status(5) + 1
      Case Else
          data_status(6) = data_status(6) + 1
      End Select
      
    Next z
 
 
 
 
 
Next t


For Y = 0 To 11
  grafica3.Row = Y + 1
  grafica3.Column = 1
  grafica3.Data = data_sold(Y, 0)
  
  grafica3.Row = Y + 1
  grafica3.Column = 2
  grafica3.Data = data_sold(Y, 1)
  
Next Y


For z = 1 To 7
 grafica2.Column = z
 grafica2.Data = data_status(z - 1)
 lbl_cant_status(z - 1).Caption = data_status(z - 1)
Next z


End Sub

Public Sub calcula_datos_mes()
On Error Resume Next
mensaje.Visible = True
mensaje.Refresh

grafica1.RowCount = 1


Dim data_status(7) As Integer, data_sold(12, 2) As Integer
Erase data_status
Erase data_sold

'For t = 0 To 11
' grafica1.Column = t + 1
' grafica1.Data = t + 1
 
'Next t





If mes_o_dia = 0 Then

  Select Case mes_elegido
  Case 1
     date1$ = "01/01/" + Format(year1, "0000")
     date2$ = "01/31/" + Format(year1, "0000")
  Case 2
     res = (year1 / 4) - Int(year1 / 4)
     If res = 0 Then
        date1$ = "02/01/" + Format(year1, "0000")
        date2$ = "02/29/" + Format(year1, "0000")
     Else
        date1$ = "02/01/" + Format(year1, "0000")
        date2$ = "02/28/" + Format(year1, "0000")
     End If
        
  
  Case 3
     date1$ = "03/01/" + Format(year1, "0000")
     date2$ = "03/31/" + Format(year1, "0000")
  Case 4
     date1$ = "04/01/" + Format(year1, "0000")
     date2$ = "04/30/" + Format(year1, "0000")
  Case 5
     date1$ = "05/01/" + Format(year1, "0000")
     date2$ = "05/31/" + Format(year1, "0000")
  Case 6
     date1$ = "06/01/" + Format(year1, "0000")
     date2$ = "06/30/" + Format(year1, "0000")
  Case 7
     date1$ = "07/01/" + Format(year1, "0000")
     date2$ = "07/31/" + Format(year1, "0000")
  Case 8
     date1$ = "08/01/" + Format(year1, "0000")
     date2$ = "08/31/" + Format(year1, "0000")
  Case 9
     date1$ = "09/01/" + Format(year1, "0000")
     date2$ = "09/30/" + Format(year1, "0000")
  Case 10
     date1$ = "10/01/" + Format(year1, "0000")
     date2$ = "10/31/" + Format(year1, "0000")
  Case 11
     date1$ = "11/01/" + Format(year1, "0000")
     date2$ = "11/30/" + Format(year1, "0000")
  Case 12
     date1$ = "12/01/" + Format(year1, "0000")
     date2$ = "12/31/" + Format(year1, "0000")
  End Select
     
Else
    ' corregir aqui
    date1$ = Format(mes_elegido, "00") + "/" + cbodia.List(cbodia.ListIndex) + "/" + Format(year1, "0000")
    date2$ = date1$
     
End If
     
     
     
     
     
     
     
     

  If regional <= 0 Then
                                   
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              
              If f1$ = "" Then
                Exit Sub
              End If
              
              If tipo_agente = 0 Then   ' todos los agentes
              
                If callcenter1$ = "N" Then
                   sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE oficina='" + oficina_autorizada$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                Else
                   sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE  (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                End If
                
              Else   ' solo un agente
              
                If callcenter1$ = "N" Then
                   sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE csr='" + txtagente.Text + "' and oficina='" + oficina_autorizada$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                Else
                   sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE csr='" + txtagente.Text + "' and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                End If
                
              
              
              End If
    
    
  Else
   ' si es un regional entonces            region_zona
   ' ********************************************************************************************************
                            
              f1$ = "convert(datetime, '" + Format(date1$, "mm/dd/yyyy") + "')"
              f2$ = "convert(datetime, '" + Format(date2$, "mm/dd/yyyy") + "')"
              
              If f2$ = "" Then f2$ = f1$
              If f1$ = "" Then
                Exit Sub
              End If
                            
              'If callcenter1$ = "N" Then
                 
                 ' selecciona oficina individual
                 
               If tipo_agente = 0 Then   ' todos los agentes
                 
                 If region_parcial = 0 Then
                           sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE region='" + Format(region_zona, "#0") + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                 Else
                           oficina_parcial$ = matriz$(region_parcial, 0)
                           sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                    
                 
                 End If
                 
               Else
               
                 If region_parcial = 0 Then
                           sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE csr='" + txtagente.Text + "' and region='" + Format(region_zona, "#0") + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                 Else
                           oficina_parcial$ = matriz$(region_parcial, 0)
                           sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE csr='" + txtagente.Text + "' and oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                    
                 
                 End If
                 
               
               
               End If
                           
             ' Else
             '     If region_parcial = 0 Then
             '             sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE  (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
             '     Else
             '             oficina_parcial$ = matriz$(region_parcial, 0)
             '             sSelect = "SELECT idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + ") " + tipo_orden$
                  
             '     End If
                  
             ' End If
        
   
 ' **************************************************************************************
    
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
              'If i = 3 Then
                 ' cambia la oficina por su inicial
                 
               '  For Y = 0 To 100
                '    If RTrim(UCase(matriz_oficina$(Y, 0))) = RTrim(UCase(Rs.Fields(i).value)) Then
                 '      inicial_oficina$ = RTrim(UCase(matriz_oficina$(Y, 1)))
                  '     Exit For
                   ' End If
                ' Next Y
                 
                ' grid1.TextMatrix(j, i) = inicial_oficina$
                 
              'Else
                 If i = (columnas - 1) Then
                    If Rs.Fields(i).value = 0 Then
                       grid1.TextMatrix(j, i) = ""
                    Else
                       grid1.TextMatrix(j, i) = Rs.Fields(i).value
                    
                    End If
                 Else
                 
                    
                 
                    grid1.TextMatrix(j, i) = Rs.Fields(i).value
                 
                 End If
              'End If
              
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
 
 



' ****************************************************************************************************************

'      idcita, quote, oficina,fecha_cita1, hora_cita1,status_gral,cliente, telefono, direccion, ciudad, estado,cp, csr, quote_turborater FROM citas WHERE oficina='" + oficina_parcial$ + "'  and (fecha_cita1 between " + f1$ + " and " + f2$ + " or fecha_cita2 between " + f1$ + " and " + f2$ + " or fecha_cita3 between " + f1$ + " and " + f2$ + ") order by quote asc"
                  

' asigna anchos de columnas
    grid1.ColWidth(0) = 900   ' idcita
    grid1.ColWidth(1) = 1000  ' quote
    grid1.ColWidth(2) = 1800 ' oficina
    grid1.ColWidth(3) = 1250 ' fecha appointment
    grid1.ColWidth(4) = 820  '  hora cita
    grid1.ColWidth(5) = 1800  ' status
    grid1.ColWidth(6) = 3000  ' customer
    grid1.ColWidth(7) = 1400   ' telefono
    grid1.ColWidth(8) = 3200  ' direccion
    grid1.ColWidth(9) = 2000  ' ciudad
    grid1.ColWidth(10) = 600  ' estado
    grid1.ColWidth(11) = 700  ' zip
    grid1.ColWidth(12) = 800  '  csr
    grid1.ColWidth(13) = 1000  ' quote turborater
    
    
    ' cambia los titulos del GRID
    grid1.Row = 0
    
    grid1.Col = 0
    grid1.Text = ""
    
    grid1.RowHeight(0) = 600
    grid1.ColAlignment(0) = 4   ' 1=izq   4=centro  7=derecha
    grid1.ColAlignment(1) = 4
    
    For Y = 2 To 13
       grid1.ColAlignment(Y) = 1
    Next Y
    
    
    
    grid1.Col = 1
    grid1.Text = "Quote"
    
    grid1.Col = 2
    grid1.Text = "Office"
    
    grid1.Col = 3
    grid1.Text = "Date"
    
    grid1.Col = 4
    grid1.Text = "Time"
    
    grid1.Col = 5
    grid1.Text = "Status"
    
    grid1.Col = 6
    grid1.Text = "Customer"
    
    grid1.Col = 7
    grid1.Text = "Phone"
    
    grid1.Col = 8
    grid1.Text = "Address"
    
    grid1.Col = 9
    grid1.Text = "City"
    
    grid1.Col = 10
    grid1.Text = "State"
    
    grid1.Col = 11
    grid1.Text = "Zip"
    
    grid1.Col = 12
    grid1.Text = "CSR"
    
    grid1.Col = 13
    grid1.Text = "Quote TR"
    
mensaje.Visible = False


End Sub

Private Sub txtagente_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
  calcula_datos_mes

End If


End Sub


