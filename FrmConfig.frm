VERSION 5.00
Begin VB.Form FrmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail account"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "FrmConfig.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox btnsetup 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Text            =   "25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CheckBox chkAut 
      Caption         =   "This account requires authentication"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   5655
   End
   Begin VB.CheckBox chkSSL 
      Caption         =   "Use SSL - Some servers require this option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   6375
   End
   Begin VB.ComboBox cboServ 
      Height          =   315
      ItemData        =   "FrmConfig.frx":000C
      Left            =   5400
      List            =   "FrmConfig.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   1720
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1125
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server SMTP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   555
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private mcIni                       As clsIni               ' guardar los datos en un archivo .ini
Dim INI_PATH                        As String

Private Sub btnsetup_Click()
On Error Resume Next

txtServer.Text = "smtp.office365.com"
txtPort.Text = "25"
chkAut.value = 1
chkSSL.value = 1
txtUser.Text = "hnavarro@justautoins.com     "
txtPassword.Text = "Jessicab28$"

End Sub

Private Sub CmdAceptar_Click()
On Error Resume Next
    actualiza_registro


    With mcIni
        Call .writeValue(INI_PATH, "datos", "servicio Mail", cboServ.ListIndex)
        Call .writeValue(INI_PATH, "datos", "servidor", txtServer.Text)
        Call .writeValue(INI_PATH, "datos", "usuario", txtUser.Text)
        Call .writeValue(INI_PATH, "datos", "password", .Encriptar(App.EXEName, txtPassword.Text, 1))
        Call .writeValue(INI_PATH, "datos", "puerto", txtPort.Text)
        Call .writeValue(INI_PATH, "datos", "ssl", chkSSL.value)
        Call .writeValue(INI_PATH, "datos", "Aut", chkAut.value)
    End With
    
    
    
    
    
    base.Close
    
    If transfiere$ = "888" Then
      Hide
      base.Close
     
      Unload Me
      Exit Sub
    End If
    
    
    Unload Me
End Sub

Public Sub actualiza_registro()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
    
    ' DETECTA SI ESTA CREADO EL REGISTRO

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT usuario From correo where idcorreo='1'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    usuario_correo$ = Rs(0)
    
                         
    Rs.Close
    
    
    
    
    
       
       Set Rs = New ADODB.Recordset
    
    If usuario_correo$ <> "" Then
                             
       sSelect = "update correo set servicio='2', server='" + txtServer.Text + "', port='" + txtPort.Text + "', aut='" + Format(chkAut.value, "0") + _
    "', SSL='" + Format(chkSSL.value, "0") + "', usuario='" + txtUser.Text + "', password='" + txtPassword.Text + "' where idcorreo='1'"
    
    Else
    
    sSelect = "INSERT INTO CORREO (Idcorreo, servicio, server, port, aut, ssl, usuario, password)  VALUES ('1" + _
    "', '2', '" + txtServer.Text + "', '" + txtPort.Text + "', '" + Format(chkAut.value, "0") + "', '" + Format(chkSSL.value, "0") + _
    "', '" + txtUser.Text + "', '" + txtPassword.Text + "')"
    
    End If
    
    
       Rs.Open sSelect, base, adOpenUnspecified
    
       Rs.Close
    
    
    
End Sub
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
'Me.BackColor = Form1.BackColor
'chkAut.BackColor = Form1.BackColor
'chkSSL.BackColor = Form1.BackColor

If transfiere$ = "888" Then Exit Sub


Conecta_SQL



 Dim sSelect As String
 Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
 Dim Rs As ADODB.Recordset
    
   
    
cboServ.ListIndex = 2
    
    
 
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT server From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    servidor$ = Rs(0)
                         
    Rs.Close
    
    txtServer.Text = servidor$





    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT port From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    puerto$ = Rs(0)
                         
    Rs.Close

    txtPort.Text = puerto$
         
         
         
         
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT usuario From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    usuario$ = Rs(0)
                         
    Rs.Close
         
    txtUser.Text = usuario$
    
    
    
    
    
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT password From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    password$ = Rs(0)
                         
    Rs.Close
    
    txtPassword.Text = password$
    
    
    
    
    
    
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT aut From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    auten$ = Rs(0)
                         
    Rs.Close
    
    
       
         chkAut.value = Val(auten$)




    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT ssl From correo where idcorreo='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ssl$ = Rs(0)
                         
    Rs.Close
    
     chkSSL.value = Val(ssl$)





    Set mcIni = New clsIni
    ' INI_PATH = App.Path & "\config.ini"
   INI_PATH = "c:\discrepancy\config.ini"



cboServ.ListIndex = 2
cboServ.Enabled = False
cboServ.Visible = False


    With mcIni
        Call .writeValue(INI_PATH, "datos", "servicio Mail", cboServ.ListIndex)
        Call .writeValue(INI_PATH, "datos", "servidor", txtServer.Text)
        Call .writeValue(INI_PATH, "datos", "usuario", txtUser.Text)
        Call .writeValue(INI_PATH, "datos", "password", .Encriptar(App.EXEName, txtPassword.Text, 1))
        Call .writeValue(INI_PATH, "datos", "puerto", txtPort.Text)
        Call .writeValue(INI_PATH, "datos", "ssl", chkSSL.value)
        Call .writeValue(INI_PATH, "datos", "Aut", chkAut.value)
    End With






    
    With mcIni
         cboServ.ListIndex = .getValue(INI_PATH, "datos", "servicio Mail", -1)
         txtServer.Text = .getValue(INI_PATH, "datos", "servidor")
         txtPort.Text = .getValue(INI_PATH, "datos", "puerto")
         txtUser.Text = .getValue(INI_PATH, "datos", "usuario")
         txtPassword.Text = .Encriptar(App.EXEName, .getValue(INI_PATH, "datos", "password"), 2)
         chkSSL.value = .getValue(INI_PATH, "datos", "ssl", 0)
         chkAut.value = .getValue(INI_PATH, "datos", "Aut", 0)
     End With
     
     
     If transfiere$ = "777" Then
        CmdAceptar_Click
        
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
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Combobox para Indicar el servicio Smtp de mail a utilizar ( Yahoo, Gmail y otros ...)
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cboServ_Click()
    Dim idMail  As String
    
    
    If Me.Visible = False Then
        Exit Sub
    End If
    
    With cboServ
        Select Case .ListIndex
            ' Yahoo
            Case 0
                txtPort.Text = "465"
                chkAut.value = 1
                chkSSL.value = 1
                txtServer.Text = "smtp.mail.yahoo.com"
                
                idMail = InputBox("Ingrese el Id de su cuenta de yahoo. Por ejemplo si su cuenta es 'maria@yahoo.com', puede ingresar 'maria@yahoo.com' , o solo 'maria'")
                
                If idMail <> "" Then
                    'txtFrom.Text = idMail
                    txtUser.Text = idMail
                    MsgBox "Para poder utilizar el acceso pop y Smtp de Yahoo, deberá estar activada la opción 'Acceso web y Pop', desde las opciones generales de la cuenta de Yahoo", vbInformation
                End If
            ' Gmail
            Case 1
                txtPort.Text = "465"
                chkAut.value = 1
                chkSSL.value = 1
                txtServer.Text = "smtp.gmail.com"
                
                idMail = InputBox("Ingrese el Id de su cuenta de Gmail. Por ejemplo si su cuenta es 'maria@gmail.com', puede ingresar 'maria@gmail.com' , o solo 'maria'")
                If idMail <> "" Then
                    'txtFrom.Text = idMail
                    txtUser.Text = idMail
                End If
            ' otro
            Case 2
                chkAut.value = 1
                chkSSL.value = 0
                txtServer.Text = ""
                txtPort.Text = "25"
                txtPassword.Text = ""
                'txtFrom.Text = ""
                txtUser.Text = ""
        End Select
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mcIni = Nothing
End Sub

Private Sub Image1_Click()
On Error Resume Next
txtPassword.PasswordChar = ""
End Sub


