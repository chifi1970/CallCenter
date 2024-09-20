VERSION 5.00
Begin VB.Form forma_cambia_password 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btnok 
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
   Begin VB.TextBox txt2 
      BackColor       =   &H00E0E0E0&
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
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00E0E0E0&
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
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin Project1.lvButtons_H btncancel 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   2
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
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type the password:"
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the new password:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "forma_cambia_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancel_Click()
On Error Resume Next
password$ = ""
Unload Me
End Sub

Private Sub btnok_Click()
On Error Resume Next
If txt1.Text <> txt2.Text Then
   MsgBox "The password does not match", 16, "Attention"
   txt1.SetFocus
   Exit Sub
End If


' Para la cadena de selección
       Dim sSelect As String
       Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
       Dim Rs As ADODB.Recordset
    
        
       Set Rs = New ADODB.Recordset
                         
       sSelect = "update employees set contrasena='" + txt1.Text + "' where login='" + password$ + "'"
                     
       Rs.Open sSelect, base, adOpenUnspecified
       Rs.Close
    
       password$ = ""

Unload Me


End Sub
