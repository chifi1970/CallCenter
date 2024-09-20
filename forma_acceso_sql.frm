VERSION 5.00
Begin VB.Form forma_acceso 
   BackColor       =   &H00C1CEF9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "forma_acceso_sql.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btnok 
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      Caption         =   "OK"
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   2520
      Picture         =   "forma_acceso_sql.frx":3AFA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "forma_acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnok_Click()
On Error Resume Next
password$ = RTrim(txt1.Text)
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - (Height + 4500)) / 2


End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
  btnok_Click
End If
End Sub
