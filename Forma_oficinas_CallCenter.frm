VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form forma_oficinas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Appointment Scheduler"
   ClientHeight    =   10590
   ClientLeft      =   -40770
   ClientTop       =   -3780
   ClientWidth     =   18105
   Icon            =   "Forma_oficinas_CallCenter.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   18105
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   4800
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00F7FCF5&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   13155
      TabIndex        =   10
      Top             =   1800
      Width           =   13215
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   495
         Left            =   6840
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F7FCF5&
         BorderStyle     =   0  'None
         Caption         =   "offices by region"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   11040
         TabIndex        =   28
         Top             =   120
         Width           =   1815
         Begin VB.ListBox lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00F7FCF5&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox cboregiones 
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
            Left            =   800
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Image3 
            Height          =   375
            Left            =   240
            Picture         =   "Forma_oficinas_CallCenter.frx":3336E
            Stretch         =   -1  'True
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offices by region"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   75
            Width           =   1425
         End
      End
      Begin VB.ComboBox cboregion 
         Height          =   315
         Left            =   6840
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   2280
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   3735
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   6588
         _Version        =   393216
         BackColorBkg    =   16252149
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
      Begin VB.TextBox txtemail6 
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
         Left            =   2760
         MaxLength       =   80
         TabIndex        =   9
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox txtreport_pwd6 
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
         Left            =   1080
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtmanager6 
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
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtzip6 
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
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtestado6 
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
         TabIndex        =   5
         Top             =   1580
         Width           =   495
      End
      Begin VB.TextBox txtciudad6 
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
         TabIndex        =   4
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtdireccion6 
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
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtphone6 
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
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   2
         Top             =   860
         Width           =   2415
      End
      Begin VB.TextBox txtshort_name6 
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
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   1
         Top             =   860
         Width           =   1095
      End
      Begin VB.TextBox txtoficina6 
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
         MaxLength       =   30
         TabIndex        =   0
         Top             =   840
         Width           =   2655
      End
      Begin Project1.lvButtons_H btnnew6 
         Height          =   855
         Left            =   8760
         TabIndex        =   39
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
         Image           =   "Forma_oficinas_CallCenter.frx":337B0
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnsave6 
         Height          =   855
         Left            =   8760
         TabIndex        =   40
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
         Image           =   "Forma_oficinas_CallCenter.frx":35B30
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnborra6 
         Height          =   855
         Left            =   8760
         TabIndex        =   41
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
         Image           =   "Forma_oficinas_CallCenter.frx":37833
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   26
         Top             =   2040
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   7800
         Picture         =   "Forma_oficinas_CallCenter.frx":39478
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Pwd:"
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
         Left            =   1080
         TabIndex        =   21
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office E-mail:"
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
         Left            =   2760
         TabIndex        =   20
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manager:"
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
         Width           =   690
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
         Index           =   23
         Left            =   4200
         TabIndex        =   14
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short name:"
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
         Index           =   22
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office name:"
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
         TabIndex        =   12
         Top             =   600
         Width           =   930
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Offices"
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
         TabIndex        =   11
         Top             =   120
         Width           =   8295
      End
   End
   Begin Project1.lvButtons_H btncallcenter 
      Height          =   1335
      Left            =   240
      TabIndex        =   33
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
      Image           =   "Forma_oficinas_CallCenter.frx":39EAC
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnvendor 
      Height          =   1335
      Left            =   1560
      TabIndex        =   34
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
      Image           =   "Forma_oficinas_CallCenter.frx":3B923
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnReports 
      Height          =   1335
      Left            =   2880
      TabIndex        =   35
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
      Image           =   "Forma_oficinas_CallCenter.frx":3EDFB
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnemployees 
      Height          =   1335
      Left            =   4200
      TabIndex        =   36
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
      Image           =   "Forma_oficinas_CallCenter.frx":41397
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnOffices 
      Height          =   1335
      Left            =   5520
      TabIndex        =   37
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
      Image           =   "Forma_oficinas_CallCenter.frx":43628
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnexit 
      Height          =   1335
      Left            =   11880
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_oficinas_CallCenter.frx":458B0
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   3120
      Picture         =   "Forma_oficinas_CallCenter.frx":47A9F
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
      Left            =   2160
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
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   -720
      Top             =   -120
      Width           =   17895
   End
End
Attribute VB_Name = "forma_oficinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer
Dim id_oficina As Integer, seg As Integer

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)



Public Sub Agrega_registro()
On Error Resume Next
  
  
 
' revisa el numero de id disponible


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idoficina From OFICINA ORDER BY idoficina DESC;"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ultimo_id = Rs(0)
    
    
    'If Err.Number <> 0 Then
    '             MsgBox "Error # " & Str(Err.Number) & " fue generado por " & Err.Source & Chr(13) & Err.Description
    'End If
                         
    Rs.Close
                         
                         
                         
                         
    ' inserta el registro
    
                         
    sSelect = "INSERT INTO oficina (Idoficina, nombre, abreviatura, telefono, direccion, ciudad, estado, cp, manager, contrasena, correo, region)  VALUES ('" + Format(ultimo_id + 1, "####0") + "', '" + LTrim(UCase(txtoficina6.Text)) + "', '" + LTrim(UCase(txtshort_name6.Text)) + "', '" + LTrim(txtphone6.Text) + "', '" + LTrim(txtdireccion6.Text) + "', '" + LTrim(UCase(txtciudad6.Text)) + "', '" + LTrim(UCase(txtestado6.Text)) + "', '" + LTrim(txtzip6.Text) + "', '" + LTrim(UCase(txtmanager6.Text)) + "', '" + LTrim(txtreport_pwd6.Text) + "', '" + LTrim(txtemail6.Text) + "', '" + cboregion.List(cboregion.ListIndex) + "')"
   
                      
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


Private Sub btnborra6_Click()
On Error Resume Next
  
  
If id_oficina = 0 Then
  MsgBox "Select the record you want to delete", 64, "Attention"
  Exit Sub
End If
  
Elimina_registro


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

Private Sub btnexit_Click()
On Error Resume Next
base.Close
End
End Sub















Private Sub btnnew6_Click()
On Error Resume Next
limpia_campos
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
  
If txtoficina6.Text = "" Then
   MsgBox "You need to type the office name", 64, "Attention"
   Exit Sub
End If

If txtshort_name6.Text = "" Then
   MsgBox "You need to type the short name", 64, "Attention"
   Exit Sub
End If

If txtphone6.Text = "" Then
   MsgBox "You need to type the phone number", 64, "Attention"
   Exit Sub
End If

If txtdireccion6.Text = "" Then
   MsgBox "You need to type the address", 64, "Attention"
   Exit Sub
End If

If txtciudad6.Text = "" Then
   MsgBox "You need to type the city", 64, "Attention"
   Exit Sub
End If

If txtestado6.Text = "" Then
   MsgBox "You need to type the state", 64, "Attention"
   Exit Sub
End If

If txtzip6.Text = "" Then
   MsgBox "You need to type the zip code", 64, "Attention"
   Exit Sub
End If

If txtmanager6.Text = "" Then
   MsgBox "You need to type the manager", 64, "Attention"
   Exit Sub
End If


If txtreport_pwd6.Text = "" Then
   MsgBox "You need to type the password", 64, "Attention"
   Exit Sub
End If

If txtemail6.Text = "" Then
   MsgBox "You need to type the email", 64, "Attention"
   Exit Sub
End If




' revisa is existe el numero de id


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT abreviatura From OFICINA where abreviatura='" + txtshort_name6.Text + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    n$ = Rs(0)
    
    If UCase(n$) = UCase(txtshort_name6.Text) Then
      X = 1
    Else
      X = 2
    End If
   
    
                         
    Rs.Close
    
    
     
 If X = 2 Then
   Agrega_registro
 Else
   Set Rs = New ADODB.Recordset
   sSelect = "SELECT idoficina From oficina where abreviatura='" + txtshort_name6.Text + "'"
   Rs.Open sSelect, base, adOpenUnspecified
   
   id_oficina = Rs(0)
   Rs.Close
   
   R$ = MsgBox("The office named " + UCase(txtshort_name6.Text) + " already exists. Do You want to overwrite it? ", 4, "Attention")
   If R$ = "7" Then Exit Sub
 
   actualiza_registro
 End If
 


End Sub

Private Sub btnvendor_Click()
On Error Resume Next
base.Close

Load forma_vendor
forma_vendor.Show
Unload Me
End Sub

Private Sub cboregiones_Click()
On Error Resume Next
' carga el campo de usuario

If cboregiones.ListIndex = -1 Then Exit Sub

  
  
    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT idoficina, nombre FROM oficina where region=" + cboregiones.List(cboregiones.ListIndex) + " order by nombre asc"
   


    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    Rs.Open sSelect, base, adOpenStatic

    ' Permitir redimensionar las columnas
    grid2.AllowUserResizing = flexResizeColumns

    ' define cuantos registros tiene la tabla
    columnas = Rs.Fields.Count
    filas = Rs.RecordCount
'Llenar las filas

   grid2.rows = filas + 1
    grid2.cols = columnas
    
    grid2.Clear



        
       
       For i = 0 To columnas
        grid2.TextMatrix(0, i) = Rs.Fields(i).Name
       Next i
       
       contador = 0
      'Llenar las filas
       For j = 1 To filas 'comenzamos en 1 porque el encabezado no se vuelve a llenar
          contador = contador + 1
          grid2.TextMatrix(j, 0) = Format(contador, "####0")
          For i = 1 To columnas
              grid2.TextMatrix(j, i) = Rs.Fields(i).value
          Next i
              
          Rs.MoveNext 'al terminar de llenar todas las columnas brincar al siguiente registro
       Next j
    
        
    
    Rs.Close
    
lista.Clear
For t = 1 To (grid2.rows - 1)
  grid2.Row = t
  grid2.Col = 1
  lista.AddItem grid2.Text
Next t


    
    
End Sub


Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2

If administrador$ = "Y" Then
   btnvendor.Enabled = False
End If

lblfecha.Caption = Format(Now, "mm/dd/yyyy")

Dim lRet As Long
    lRet = GetWindowLong(Me.hWnd, GWL_STYLE)
    lRet = lRet And Not (WS_MAXIMIZEBOX)
    lRet = SetWindowLong(Me.hWnd, GWL_STYLE, lRet)
    DesactivarMenu Me
    
 id_oficina = 0

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
            DesignY = 940 '1024 '800
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

cboregion.Clear
cboregiones.Clear

For t = 1 To 30
  cboregion.AddItem t
  cboregiones.AddItem t
Next t

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

txtoficina6.Text = ""
txtshort_name6.Text = ""
txtphone6.Text = ""
txtdireccion6.Text = ""
txtciudad6.Text = ""
txtestado6.Text = ""
txtzip6.Text = ""
txtmanager6.Text = ""
txtreport_pwd6.Text = ""
txtemail6.Text = ""
cboregion.ListIndex = -1
id_oficina = 0

End Sub

Public Sub carga_registros()
On Error Resume Next


  
    ' Para la cadena de selección
    Dim sSelect As String
    
    ' Para una base de datos normal:
     sSelect = "SELECT idoficina, nombre, abreviatura, telefono, direccion, ciudad, estado, cp, manager, region FROM oficina order by nombre asc"
   


    
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
    grid1.ColWidth(1) = 2800
    grid1.ColWidth(2) = 1300
    grid1.ColWidth(3) = 1600
    grid1.ColWidth(4) = 3200
    grid1.ColWidth(5) = 2200
    grid1.ColWidth(6) = 800
    grid1.ColWidth(7) = 900
    grid1.ColWidth(8) = 1100
    grid1.ColWidth(9) = 800
    
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
    grid1.ColAlignment(5) = 1
    grid1.ColAlignment(6) = 1
    
    grid1.ColAlignment(7) = 4
    grid1.ColAlignment(8) = 1
    grid1.ColAlignment(9) = 1
    
    
    grid1.Col = 1
    grid1.Text = "Office"
    
    grid1.Col = 2
    grid1.Text = "Short Name"
    
    grid1.Col = 3
    grid1.Text = "Phone"
    
    grid1.Col = 4
    grid1.Text = "Address"
    
    grid1.Col = 5
    grid1.Text = "City"
    
    grid1.Col = 6
    grid1.Text = "State"
    
    grid1.Col = 7
    grid1.Text = "ZIP"
    
    grid1.Col = 8
    grid1.Text = "Manager"
    
    grid1.Col = 9
    grid1.Text = "Region"
    
    
    
    
    
    
    
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
    
                         
    sSelect = "update oficina set nombre='" + LTrim(UCase(txtoficina6.Text)) + "', abreviatura='" + LTrim(txtshort_name6.Text) + "', telefono='" + LTrim(txtphone6.Text) + "', direccion='" + LTrim(txtdireccion6.Text) + "', ciudad='" + LTrim(UCase(txtciudad6.Text)) + "', estado='" + LTrim(UCase(txtestado6.Text)) + "', cp='" + LTrim(txtzip6.Text) + "', manager='" + LTrim(UCase(txtmanager6.Text)) + "', contrasena='" + LTrim(txtreport_pwd6.Text) + "', correo='" + LTrim(txtemail6.Text) + "', region='" + cboregion.List(cboregion.ListIndex) + "' where idoficina='" + Format(id_oficina, "#####0") + "'"
    
      
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
    
    limpia_campos
    
    carga_registros
    
    
    
End Sub

Private Sub grid1_Click()
On Error Resume Next

  fila = grid1.Row
  
  'grid1.Col = 0
  'id_oficina = Val(grid1.Text)
  
  grid1.Col = 1
  txtoficina6.Text = grid1.Text
  
  grid1.Col = 2
  txtshort_name6.Text = grid1.Text
  
  grid1.Col = 3
  txtphone6.Text = grid1.Text
  
  grid1.Col = 4
  txtdireccion6.Text = grid1.Text
  
  grid1.Col = 5
  txtciudad6.Text = grid1.Text
  
  grid1.Col = 6
  txtestado6.Text = grid1.Text
  
  grid1.Col = 7
  txtzip6.Text = grid1.Text
  
  grid1.Col = 8
  txtmanager6.Text = grid1.Text
  
  
  
  ' ************************************************************
  ' carga el ID


    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT idoficina From OFICINA where abreviatura='" + txtshort_name6.Text + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_oficina = Rs(0)
            
                         
    Rs.Close
    
 
 
 
 
   
  ' ************************************************************
  ' carga el campo de password


    ' Para la cadena de selección
   
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT contrasena From OFICINA where idoficina='" + Format(id_oficina, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtreport_pwd6.Text = Rs(0)
            
                         
    Rs.Close
    
  
  ' ************************************************************
  ' carga el email
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT correo From OFICINA where idoficina='" + Format(id_oficina, "######0") + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    txtemail6.Text = Rs(0)
            
                         
    Rs.Close
  
   ' ************************************************************
   ' region
   ' El recordset para acceder a los datos
    
    Set Rs = New ADODB.Recordset

    cboregion.ListIndex = -1
    sSelect = "SELECT region From oficina where idoficina='" + Format(id_oficina, "######0") + "'"
    

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
    
                         
    sSelect = "delete from oficina  where idoficina='" + Format(id_oficina, "#####0") + "'"
    
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
     limpia_campos
    
    carga_registros
    
End Sub

Private Sub Timer1_Timer()
seg = seg + 1
lblhora.Caption = Format(Now, "hh:mm am/pm")
End Sub


Private Sub txtzip6_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 8 Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


