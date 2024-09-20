VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form forma_employees 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Appointment Scheduler"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18105
   Icon            =   "forma_employees_callcenter.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11520
      Top             =   9720
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EBF4FE&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   13035
      TabIndex        =   11
      Top             =   1800
      Width           =   13095
      Begin VB.CommandButton btnlimpiar 
         Caption         =   "X"
         Height          =   375
         Left            =   10200
         TabIndex        =   63
         Top             =   240
         Width           =   375
      End
      Begin Project1.lvButtons_H btnbusca 
         Height          =   735
         Left            =   6840
         TabIndex        =   61
         Top             =   120
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
         Image           =   "forma_employees_callcenter.frx":3336E
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtbusca 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   62
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Full Access:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   37
         Top             =   3840
         Width           =   1335
         Begin Project1.lvButtons_H op_full 
            Height          =   495
            Index           =   0
            Left            =   500
            TabIndex        =   57
            Top             =   180
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_full 
            Height          =   495
            Index           =   1
            Left            =   860
            TabIndex        =   58
            Top             =   180
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image9 
            Height          =   480
            Left            =   80
            Picture         =   "forma_employees_callcenter.frx":34CD5
            Stretch         =   -1  'True
            Top             =   180
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Administrator:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7800
         TabIndex        =   36
         Top             =   3840
         Width           =   1335
         Begin Project1.lvButtons_H op_admin 
            Height          =   495
            Index           =   0
            Left            =   540
            TabIndex        =   59
            Top             =   180
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_admin 
            Height          =   495
            Index           =   1
            Left            =   920
            TabIndex        =   60
            Top             =   180
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image6 
            Height          =   560
            Left            =   40
            Picture         =   "forma_employees_callcenter.frx":353DE
            Stretch         =   -1  'True
            Top             =   160
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Office: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         TabIndex        =   28
         Top             =   2760
         Width           =   5295
         Begin VB.ComboBox cbooficinas 
            BackColor       =   &H00C0E0FF&
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
            Left            =   1440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   440
            Width           =   2775
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_oficina 
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   50
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image2 
            Height          =   660
            Left            =   360
            Picture         =   "forma_employees_callcenter.frx":35D77
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4815
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Manager:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   31
         Top             =   3600
         Width           =   2895
         Begin VB.ComboBox cboregion 
            BackColor       =   &H00EBF4FE&
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
            Left            =   2320
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   560
            Width           =   495
         End
         Begin VB.OptionButton op_tipo_manager 
            BackColor       =   &H00EBF4FE&
            Caption         =   "Region"
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
            Index           =   1
            Left            =   1680
            TabIndex        =   34
            Top             =   560
            Width           =   735
         End
         Begin VB.OptionButton op_tipo_manager 
            BackColor       =   &H00EBF4FE&
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
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin Project1.lvButtons_H op_manager 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_manager 
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   52
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image3 
            Height          =   780
            Left            =   900
            Picture         =   "forma_employees_callcenter.frx":3653B
            Stretch         =   -1  'True
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Call Center:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4920
         TabIndex        =   30
         Top             =   3720
         Width           =   1095
         Begin Project1.lvButtons_H op_callcenter 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_callcenter 
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   56
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image4 
            Height          =   420
            Left            =   360
            Picture         =   "forma_employees_callcenter.frx":370CC
            Stretch         =   -1  'True
            Top             =   380
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EBF4FE&
         Caption         =   "Active: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   29
         Top             =   3840
         Width           =   1095
         Begin Project1.lvButtons_H op_active 
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   53
            Top             =   200
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "Y"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin Project1.lvButtons_H op_active 
            Height          =   495
            Index           =   1
            Left            =   720
            TabIndex        =   54
            Top             =   200
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   873
            Caption         =   "N"
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   15463678
         End
         Begin VB.Image Image5 
            Height          =   375
            Left            =   80
            Picture         =   "forma_employees_callcenter.frx":37C12
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtpassword 
         BackColor       =   &H00C0FFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox txtlogin 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   20
         TabIndex        =   9
         Top             =   3000
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2175
         Left            =   240
         TabIndex        =   22
         Top             =   4800
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3836
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   15463678
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
      Begin VB.TextBox txtemail 
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
         Left            =   4560
         MaxLength       =   80
         TabIndex        =   8
         Top             =   2300
         Width           =   4575
      End
      Begin VB.TextBox txtcelular 
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
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2300
         Width           =   1695
      End
      Begin VB.TextBox txtphone 
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
         TabIndex        =   6
         Top             =   2300
         Width           =   1695
      End
      Begin VB.TextBox txtzip 
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
         Left            =   7800
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1580
         Width           =   615
      End
      Begin VB.TextBox txtestado 
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
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtciudad 
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
         Left            =   4680
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtdireccion 
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
         TabIndex        =   2
         Top             =   1580
         Width           =   4335
      End
      Begin VB.TextBox txtapellido 
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
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtnombre 
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
         Height          =   375
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         Top             =   840
         Width           =   2655
      End
      Begin Project1.lvButtons_H btnnew6 
         Height          =   855
         Left            =   11880
         TabIndex        =   44
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
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
         Image           =   "forma_employees_callcenter.frx":385AA
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnsave6 
         Height          =   855
         Left            =   11880
         TabIndex        =   45
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
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
         Image           =   "forma_employees_callcenter.frx":3A92A
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnborra6 
         Height          =   855
         Left            =   11880
         TabIndex        =   46
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1508
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_employees_callcenter.frx":3C62D
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnajusta_base 
         Height          =   975
         Left            =   10800
         TabIndex        =   47
         Top             =   3240
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
         Image           =   "forma_employees_callcenter.frx":3E272
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnsetup 
         Height          =   975
         Left            =   11760
         TabIndex        =   48
         Top             =   3240
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
         Image           =   "forma_employees_callcenter.frx":3FD92
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.Image Image11 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image10 
         Height          =   735
         Left            =   11040
         Picture         =   "forma_employees_callcenter.frx":41CA4
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1695
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   11880
         Picture         =   "forma_employees_callcenter.frx":420E6
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login:"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Phone:"
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
         Index           =   30
         Left            =   2040
         TabIndex        =   21
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
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
         Index           =   29
         Left            =   4560
         TabIndex        =   20
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Index           =   28
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   510
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
         Index           =   27
         Left            =   7800
         TabIndex        =   18
         Top             =   1320
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
         Index           =   26
         Left            =   7200
         TabIndex        =   17
         Top             =   1320
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
         Index           =   25
         Left            =   4680
         TabIndex        =   16
         Top             =   1320
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
         Index           =   24
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Index           =   21
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Employees"
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
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   8295
      End
   End
   Begin Project1.lvButtons_H btncallcenter 
      Height          =   1335
      Left            =   240
      TabIndex        =   38
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
      Image           =   "forma_employees_callcenter.frx":42528
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnvendor 
      Height          =   1335
      Left            =   1560
      TabIndex        =   39
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
      Image           =   "forma_employees_callcenter.frx":43F9F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnReports 
      Height          =   1335
      Left            =   2880
      TabIndex        =   40
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
      Image           =   "forma_employees_callcenter.frx":47477
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnemployees 
      Height          =   1335
      Left            =   4200
      TabIndex        =   41
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
      Image           =   "forma_employees_callcenter.frx":49A13
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnOffices 
      Height          =   1335
      Left            =   5520
      TabIndex        =   42
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
      Image           =   "forma_employees_callcenter.frx":4BCA4
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnexit 
      Height          =   1335
      Left            =   11880
      TabIndex        =   43
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
      Image           =   "forma_employees_callcenter.frx":4DF2C
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Image Image7 
      Height          =   735
      Left            =   3240
      Picture         =   "forma_employees_callcenter.frx":5011B
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   1215
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
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2040
      TabIndex        =   23
      Top             =   9240
      Width           =   2415
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
      TabIndex        =   25
      Top             =   11880
      Width           =   1575
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E6E6E6&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00F3F3F3&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   9120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   -840
      Top             =   -120
      Width           =   18015
   End
End
Attribute VB_Name = "forma_employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer
Dim id_employee As Integer, seg As Integer, oficina$, activo$, CallCenter$, manager$, tipo_manager As Integer, region As Integer, admon$, fullaccess$

Dim matriz_oficina$(100, 4)
Dim matriz$(100, 2)


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)



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
    
    
       total_reg = cbooficinas.ListCount ' - 1
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
Public Sub Agrega_registro()
On Error Resume Next
  
  
 
' revisa el numero de id disponible


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idemployee From employees ORDER BY idemployee DESC;"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ultimo_id = Rs(0)
    
    
 '   If Err.Number <> 0 Then
    '             MsgBox "Error # " & Str(Err.Number) & " fue generado por " & Err.Source & Chr(13) & Err.Description
 '   End If
                         
    Rs.Close
                         
                         
                         
                         
    ' inserta el registro
inserta:
                         
                         
    If manager$ = "N" Then
       tipo_manager = -1
    End If
                         
                         
    sSelect = "INSERT INTO employees (Idemployee, nombre, apellido, telefono, cellular, direccion, ciudad, estado, cp, correo, login, contrasena, oficina, activo, callcenter, manager, region, tipo_manager, nombre_oficina, admin, accesoTotal)  VALUES ('" + _
    Format(ultimo_id + 1, "####0") + "', '" + txtnombre.Text + "', '" + txtapellido.Text + "', '" + txtphone.Text + "', '" + txtcelular.Text + "', '" + txtdireccion.Text + _
    "', '" + txtciudad.Text + "', '" + txtestado.Text + "', '" + txtzip.Text + "', '" + txtemail.Text + "', '" + UCase(txtlogin.Text) + "', '" + txtPassword.Text + "', '" + oficina$ + _
    "', '" + activo$ + "', '" + CallCenter$ + "', '" + manager$ + "', '" + Format(region, "0") + "', '" + Format(tipo_manager, "0") + "', '" + cbooficinas.List(cbooficinas.ListIndex) + "', '" + admon$ + "', '" + fullaccess$ + "')"
   
   
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
            
    limpia_campos
    
    carga_registros
    
    
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


Private Sub btnajusta_base_Click()
On Error Resume Next

year1 = Format(Now, "yyyy")
    
    
R$ = MsgBox("Do you want to delete all the registers before " + Format(year1 - 1, "0000") + "?", 4, "Attention")
If R$ = "7" Then Exit Sub

elimina_registros_antiguos


End Sub

Private Sub btnborra6_Click()
On Error Resume Next
  
  
If id_employee = 0 Then
  MsgBox "Select the record you want to delete", 64, "Attention"
  Exit Sub
End If
  
Elimina_registro


End Sub

Private Sub btnbusca_Click()
On Error Resume Next


    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT idemployee, nombre, apellido, telefono, login, oficina, activo, callcenter, manager, admin, accesototal FROM employees WHERE nombre like '%" + UCase(txtbusca.Text) + "%' or apellido like '%" + UCase(txtbusca.Text) + "%' or login like '%" + UCase(txtbusca.Text) + "%' order by nombre asc"
     



    
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
    grid1.ColWidth(1) = 2000
    grid1.ColWidth(2) = 2000
    grid1.ColWidth(3) = 1800
    grid1.ColWidth(4) = 1200
    grid1.ColWidth(5) = 800
    grid1.ColWidth(6) = 800
    
    ' cambia los titulos del GRID
    grid1.Row = 0
    
    grid1.Col = 0
    grid1.Text = ""
    grid1.RowHeight(0) = 600
    grid1.ColAlignment(0) = 4   ' 1=izq   4=centro  7=derecha
    grid1.ColAlignment(1) = 1
    grid1.ColAlignment(2) = 1
    grid1.ColAlignment(3) = 1
    grid1.ColAlignment(4) = 1
    grid1.ColAlignment(5) = 4
    grid1.ColAlignment(6) = 4
    grid1.ColAlignment(7) = 4
    grid1.ColAlignment(8) = 4
    grid1.ColAlignment(9) = 4
    grid1.ColAlignment(10) = 4
    
    grid1.Col = 1
    grid1.Text = "First Name"
    
    grid1.Col = 2
    grid1.Text = "Last name"
    
    grid1.Col = 3
    grid1.Text = "Phone"
    
    grid1.Col = 4
    grid1.Text = "Login"
    
    grid1.Col = 5
    grid1.Text = "Office"
    
    grid1.Col = 6
    grid1.Text = "Active"
    
    grid1.Col = 7
    grid1.Text = "Call Center"
    
    grid1.Col = 8
    grid1.Text = "Manager"
    
    grid1.Col = 9
    grid1.Text = "Admin"
    
    grid1.Col = 10
    grid1.Text = "Full Access"
    
    




' =================================================


 
    
End Sub

Private Sub btncallcenter_Click()
On Error Resume Next
base.Close

Load forma_main
forma_main.Show
Unload Me

End Sub



Private Sub btnexit_Click()
On Error Resume Next
base.Close
End
End Sub















Private Sub btnlimpiar_Click()
On Error Resume Next
txtbusca.Text = ""

End Sub

Private Sub btnnew6_Click()
On Error Resume Next
limpia_campos
End Sub

Private Sub btnOffices_Click()
On Error Resume Next
base.Close

Load forma_oficinas
forma_oficinas.Show
Unload Me
End Sub

Private Sub btnReports_Click()
On Error Resume Next
base.Close

Load forma_graficas
forma_graficas.Show
Unload Me
End Sub

Private Sub btnsave6_Click()
On Error Resume Next
  
If txtnombre.Text = "" Then
   MsgBox "You need to type the First Name", 64, "Attention"
   Exit Sub
End If

If txtapellido.Text = "" Then
   MsgBox "You need to type the Last Name", 64, "Attention"
   Exit Sub
End If


If txtlogin.Text = "" Then
   MsgBox "You need to type the login", 64, "Attention"
   Exit Sub
End If


If txtPassword.Text = "" Then
   MsgBox "You need to type the password", 64, "Attention"
   Exit Sub
End If

If oficina$ = "" Then
   MsgBox "You need to select Y/N on office", 64, "Attention"
   Exit Sub
End If

If activo$ = "" Then
   MsgBox "You need to select Y/N on active option", 64, "Attention"
   Exit Sub
End If




' revisa is existe el nombre


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT nombre From employees where nombre='" + txtnombre.Text + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    n$ = Rs(0)
    
    If UCase(n$) = UCase(txtnombre.Text) Then
      X1 = 1
    Else
      X1 = 2
    End If
   
                     
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT apellido From employees where apellido='" + txtapellido.Text + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    n$ = Rs(0)
    
    If UCase(n$) = UCase(txtapellido.Text) Then
      X2 = 1
    Else
      X2 = 2
    End If
       
                         
    Rs.Close
    
    
    
    
    
    
     
 If X2 = 1 And X1 = 1 Then
   Set Rs = New ADODB.Recordset
   sSelect = "SELECT idemployee From employees where nombre='" + txtnombre.Text + "' and apellido='" + txtapellido.Text + "'"
   Rs.Open sSelect, base, adOpenUnspecified
   
   id_employee = Rs(0)
   Rs.Close
   
   If id_employee = 0 Then GoTo nuevo_empleado
   
   R$ = MsgBox("The employee named " + UCase(txtnombre.Text) + " " + UCase(txtapellido.Text) + " already exists. Do You want to overwrite it? ", 4, "Attention")
   If R$ = "7" Then Exit Sub
    
   actualiza_registro
   
 Else
nuevo_empleado:
   Agrega_registro
 End If
 


End Sub

Private Sub btnsetup_Click()
On Error Resume Next
Load forma_acceso
forma_acceso.Show 1

If password$ = "" Then Exit Sub

If password$ = "Tech789" Then
   Load FrmConfig
   FrmConfig.Show 1
Else
   MsgBox "Password is not valid", 16, "Access denied"
End If

End Sub

Private Sub btnvendor_Click()
On Error Resume Next
base.Close

Load forma_vendor
forma_vendor.Show
Unload Me
End Sub

Private Sub cboregion_Click()
On Error Resume Next
region = cboregion.List(cboregion.ListIndex)
End Sub


Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2

lblfecha.Caption = Format(Now, "mm/dd/yyyy")

If administrador$ = "Y" Then
   btnvendor.Enabled = False
End If


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
          DesignX = 1024 ' 1280
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

carga_oficinas
carga_registros


cboregion.Clear
For t = 1 To 30
  cboregion.AddItem t
Next t

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

txtnombre.Text = ""
txtapellido.Text = ""
txtphone.Text = ""
txtcelular.Text = ""
txtdireccion.Text = ""
txtciudad.Text = ""
txtestado.Text = ""
txtzip.Text = ""
txtemail.Text = ""
txtlogin.Text = ""
txtPassword.Text = ""
oficina$ = ""
activo$ = ""
admon$ = ""
fullaccess$ = ""

op_oficina(0).value = False
op_oficina(1).value = False

op_active(0).value = False
op_active(1).value = False

op_callcenter(0).value = False
op_callcenter(1).value = False

op_manager(0).value = False
op_manager(1).value = False

op_tipo_manager(0).value = False
op_tipo_manager(1).value = False

op_full(0).value = False
op_full(1).value = False

op_admin(0).value = False
op_admin(1).value = False


id_employee = 0


cbooficinas.ListIndex = -1
cboregion.ListIndex = -1
region = -1
tipo_manager = -1

txtnombre.SetFocus

End Sub

Public Sub carga_registros()
On Error Resume Next


  
    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT idemployee, nombre, apellido, telefono, login, oficina, activo, callcenter, manager, admin, accesototal FROM employees order by nombre asc"
   


    
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
    grid1.ColWidth(1) = 2000
    grid1.ColWidth(2) = 2000
    grid1.ColWidth(3) = 1800
    grid1.ColWidth(4) = 1200
    grid1.ColWidth(5) = 800
    grid1.ColWidth(6) = 800
    
    ' cambia los titulos del GRID
    grid1.Row = 0
    
    grid1.Col = 0
    grid1.Text = ""
    grid1.RowHeight(0) = 600
    grid1.ColAlignment(0) = 4   ' 1=izq   4=centro  7=derecha
    grid1.ColAlignment(1) = 1
    grid1.ColAlignment(2) = 1
    grid1.ColAlignment(3) = 1
    grid1.ColAlignment(4) = 1
    grid1.ColAlignment(5) = 4
    grid1.ColAlignment(6) = 4
    grid1.ColAlignment(7) = 4
    grid1.ColAlignment(8) = 4
    grid1.ColAlignment(9) = 4
    grid1.ColAlignment(10) = 4
    
    grid1.Col = 1
    grid1.Text = "First Name"
    
    grid1.Col = 2
    grid1.Text = "Last name"
    
    grid1.Col = 3
    grid1.Text = "Phone"
    
    grid1.Col = 4
    grid1.Text = "Login"
    
    grid1.Col = 5
    grid1.Text = "Office"
    
    grid1.Col = 6
    grid1.Text = "Active"
    
    grid1.Col = 7
    grid1.Text = "Call Center"
    
    grid1.Col = 8
    grid1.Text = "Manager"
    
    grid1.Col = 9
    grid1.Text = "Admin"
    
    grid1.Col = 10
    grid1.Text = "Full Access"
    
    
    
    
    
End Sub

Public Sub actualiza_registro()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

                         
    ' Modifica el registro
    
    If manager$ = "N" Then
       tipo_manager = -1
    End If
                         
    sSelect = "update employees set nombre='" + txtnombre.Text + "', apellido='" + txtapellido.Text + "', telefono='" + txtphone.Text + "', cellular='" + txtcelular.Text + _
    "', direccion='" + txtdireccion.Text + "', ciudad='" + txtciudad.Text + "', estado='" + txtestado.Text + "', cp='" + txtzip.Text + "', correo='" + txtemail.Text + _
    "', contrasena='" + txtPassword.Text + "', oficina='" + oficina$ + "', activo='" + activo$ + "', login='" + UCase(txtlogin.Text) + "', callcenter='" + CallCenter$ + _
    "', manager='" + manager$ + "', region='" + Format(region, "0") + "', tipo_manager='" + Format(tipo_manager, "0") + "', nombre_oficina='" + cbooficinas.List(cbooficinas.ListIndex) + "', admin='" + admon$ + "', Accesototal='" + fullaccess$ + "' where idemployee='" + Format(id_employee, "#####0") + "'"
    
      
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
    
    limpia_campos
    
    carga_registros
    
    
    
End Sub

Private Sub grid1_Click()
On Error Resume Next

  fila = grid1.Row
  
  'grid1.Col = 0
  'id_employee = Val(grid1.Text)
  
  grid1.Col = 1
  txtnombre.Text = grid1.Text
  
  grid1.Col = 2
  txtapellido.Text = grid1.Text
  
  grid1.Col = 3
  txtphone.Text = grid1.Text
  
  grid1.Col = 4
  txtlogin.Text = grid1.Text
  
  grid1.Col = 5
  Valor$ = grid1.Text
  
  If Valor$ = "Y" Then
    op_oficina(0).value = True
    oficina$ = "Y"
  Else
    op_oficina(1).value = True
    oficina$ = "N"
  End If
  
  
  
  
  grid1.Col = 6
  Valor$ = grid1.Text
  
    
  If Valor$ = "Y" Then
    op_active(0).value = True
    activo$ = "Y"
  Else
    op_active(1).value = True
    activo$ = "N"
  End If
  
  
  
 
 grid1.Col = 7
  Valor$ = grid1.Text
  
    
  If Valor$ = "Y" Then
    op_callcenter(0).value = True
    CallCenter$ = "Y"
  Else
    op_callcenter(1).value = True
    CallCenter$ = "N"
  End If
 
 
 
 grid1.Col = 8
  Valor$ = grid1.Text
  
    
  If Valor$ = "Y" Then
    op_manager(0).value = True
    manager$ = "Y"
  Else
    op_manager(1).value = True
    manager$ = "N"
  End If
 
 
 
 grid1.Col = 9
  Valor$ = grid1.Text
  
    
  If Valor$ = "Y" Then
    op_admin(0).value = True
    admon$ = "Y"
  Else
    op_admin(1).value = True
    admon$ = "N"
  End If
 
 
 
 ' ************************************************************
  ' carga el ID


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idemployee From employees where nombre='" + txtnombre.Text + "' and apellido='" + txtapellido.Text + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee = Rs(0)
            
                         
    Rs.Close
    
    
    
    
    
 
   
  ' ************************************************************
  ' carga el campo de password


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT contrasena From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtPassword.Text = Rs(0)
            
                         
    Rs.Close
    
  
  ' ************************************************************
  ' carga el email
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT correo From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtemail.Text = Rs(0)
            
                         
    Rs.Close
  
  
  
  ' ************************************************************
  ' carga direccion
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT direccion From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtdireccion.Text = Rs(0)
            
                         
    Rs.Close
  
  
  
  ' ************************************************************
  ' carga ciudad
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT ciudad From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtciudad.Text = Rs(0)
            
                         
    Rs.Close
  
  
  ' ************************************************************
  ' carga estado
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT estado From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtestado.Text = Rs(0)
            
                         
    Rs.Close
  
  
  ' ************************************************************
  ' carga CP
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT cp From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtzip.Text = Rs(0)
            
                         
    Rs.Close
  
  
  
  ' ************************************************************
  ' carga celular
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT cellular From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtcelular.Text = Rs(0)
            
                         
    Rs.Close
  
  
  
  ' ************************************************************
   ' nombre oficina
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

    cbooficinas.ListIndex = -1
    sSelect = "SELECT nombre_oficina From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    nom_oficina$ = Rs(0)
            
    For Y = 0 To cbooficinas.ListCount - 1
      If nom_oficina$ = cbooficinas.List(Y) Then
         cbooficinas.ListIndex = Y
         Exit For
      End If
    Next Y
            
                         
    Rs.Close
  
  
  
   ' ************************************************************
   ' region
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

    cboregion.ListIndex = -1
    sSelect = "SELECT region From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    region = Rs(0)
            
    cboregion.ListIndex = -1
    For Y = 0 To cboregion.ListCount - 1
      If region = Val(cboregion.List(Y)) Then
         cboregion.ListIndex = Y
         Exit For
      End If
    Next Y
            
                         
    Rs.Close
  
  
  ' ************************************************************
  ' carga tipo_manager
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT tipo_manager From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    tipo_manager = Rs(0)
            
                         
    Rs.Close
    
    op_tipo_manager(0).value = False
    op_tipo_manager(1).value = False
   
       
      op_tipo_manager(tipo_manager).value = True
    
  
  
  
   ' ************************************************************
  ' carga full access
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT Accesototal From employees where idemployee='" + Format(id_employee, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    fullaccess$ = Rs(0)
            
    op_full(0).value = False
    op_full(1).value = False
            
    If fullaccess$ = "" Or fullaccess$ = "N" Then
       op_full(1).value = True
    Else
       op_full(0).value = True
    End If
                         
                         
    Rs.Close
    
    
   
  
End Sub



Public Sub Elimina_registro()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

                         
    ' Elimina el registro
    
                         
    sSelect = "delete from employees  where idemployee='" + Format(id_employee, "#####0") + "'"
    
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
     limpia_campos
    
    carga_registros
    
End Sub

Private Sub Image11_DblClick()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

                         
    R$ = InputBox("Type quote number:", "Delete Quote")
    If R$ = "" Then Exit Sub
                         
    ' Elimina el registro
    
                         
    sSelect = "delete from citas  where quote='" + R$ + "'"
    
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
End Sub

Private Sub op_active_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  activo$ = "Y"
Else
  activo$ = "N"
End If

txtPassword.SetFocus

End Sub

Private Sub op_admin_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  admon$ = "Y"
Else
  admon$ = "N"
End If

txtPassword.SetFocus

End Sub

Private Sub op_callcenter_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  CallCenter$ = "Y"
Else
  CallCenter$ = "N"
End If

txtPassword.SetFocus

End Sub

Private Sub op_full_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  fullaccess$ = "Y"
Else
  fullaccess$ = "N"
End If

txtPassword.SetFocus
End Sub

Private Sub op_manager_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  manager$ = "Y"
Else
  manager$ = "N"
End If

txtPassword.SetFocus
End Sub


Private Sub op_oficina_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
  oficina$ = "Y"
Else
  oficina$ = "N"
End If

txtPassword.SetFocus

End Sub

Private Sub op_tipo_manager_Click(Index As Integer)
On Error Resume Next
tipo_manager = Index
End Sub

Private Sub Timer1_Timer()
seg = seg + 1
lblhora.Caption = Format(Now, "hh:mm am/pm")
End Sub








Private Sub txtbusca_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then btnbusca_Click
End Sub


Private Sub txtzip_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If


End Sub



Public Sub elimina_registros_antiguos()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

                         
    ' Elimina el registro
    
    year1 = Format(Now, "yyyy")
    f2$ = "12/31/" + Format(year1 - 2, "0000")

    f1$ = "convert(datetime, '" + f2$ + "')"
                         
    sSelect = "delete from citas  where fecha<=" + f1$
    
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
     limpia_campos
    
    carga_registros
End Sub
