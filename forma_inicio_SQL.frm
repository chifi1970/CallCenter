VERSION 5.00
Begin VB.Form Forma_inicio 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8190
   ControlBox      =   0   'False
   Icon            =   "forma_inicio_SQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btnclean 
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "X"
      CapAlign        =   2
      BackStyle       =   6
      Shape           =   4
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
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btncarga 
      Height          =   645
      Left            =   2880
      TabIndex        =   11
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Caption         =   "load user"
      CapAlign        =   2
      BackStyle       =   4
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
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
      cBack           =   16777215
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00E6E6E6&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   3120
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00E6E6E6&
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
      Left            =   1320
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "T9 "
      Top             =   2400
      Width           =   1935
   End
   Begin Project1.lvButtons_H btncancel 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   4
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
      cBack           =   12632256
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "v2.23"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   7560
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1965
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IT Department (MMXX)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   3795
      Left            =   -120
      Picture         =   "forma_inicio_SQL.frx":377EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8280
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   6960
      Picture         =   "forma_inicio_SQL.frx":41195
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "version 1.08"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   7680
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(T9)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   1640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe la contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1720
      Width           =   2535
   End
End
Attribute VB_Name = "Forma_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seg_end As Integer, Tipo As Integer, password$, actualiza As Integer

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)


' Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, _
    ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2


Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function
Public Sub crear_iconos()
On Error Resume Next


Dim strUserName As String
 
  strUserName = String(100, Chr$(0))
  'Get the username
  GetUserName strUserName, 100
  'strip the rest of the buffer
  strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
  
  user1$ = strUserName
  
  
  
MkDir "c:\users\" + user1$ + "\Accesos"

ruta_acceso$ = "c:\users\" + user1$ + "\Accesos\"

MkDir ruta_acceso$ + "Internet_browsers"
MkDir ruta_acceso$ + "Office_documents"
MkDir ruta_acceso$ + "My_PDF_files"
MkDir ruta_acceso$ + "My_Software"
MkDir ruta_acceso$ + "My_Pictures"
MkDir ruta_acceso$ + "Personal"
MkDir ruta_acceso$ + "Programming"







'Set obj = CreateObject("WScript.Shell")

'Set lnk = obj.DeleteFile(d$ & "\AT&T Texting.ink")


'Set obj = CreateObject("WScript.Shell")
'Set FSO = CreateObject("Scripting.FileSystemObject")
'DesktopPath = obj.SpecialFolders("Desktop")
'FSO.DeleteFile DesktopPath & "\AT&T Texting.ink"







If Dir$("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") <> "" Then
  SO = 1 ' 64 bits
  ruta_chrome$ = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
Else
  SO = 0 ' 32 bits
  ruta_chrome$ = "C:\Program Files\Google\Chrome\Application\chrome.exe"
End If




Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)


If Dir$("F:\tech\AceleraPC\ja_lae_new.ico") = "" Then
  MkDir "c:\iconos"
  ruta_iconos$ = "c:\iconos\"
Else
  ruta_iconos$ = "F:\tech\AceleraPC\"
  
End If




   Set lnk = obj.CreateShortcut(d$ & "\" & "LAE System.lnk")
   
   
   lnk.TargetPath = ruta_chrome$
   lnk.Arguments = " https://www.laesystem.com"
   lnk.Description = "LAE system"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_lae_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Sonar.lnk")
   
   
   lnk.TargetPath = ruta_chrome$
   lnk.Arguments = " dashboard.sendsonar.com/users/sign_in"
   lnk.Description = "Sonar"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_sonar_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
      
      
      
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Eversign.lnk")
   
   
   lnk.TargetPath = ruta_chrome$
   lnk.Arguments = " https://eversign.com/login"
   lnk.Description = "Eversign"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_eversign_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Appointment Scheduler.lnk")
   
   
   lnk.TargetPath = "C:\callcenter\CallCenter.exe"
   lnk.Arguments = ""
   lnk.Description = "Appointment Scheduler"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "JA_appointment_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
If Dir$("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") <> "" Then
  
  'ruta_internet$ = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
  ruta_internet$ = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Else
  
  ruta_internet$ = "C:\Program Files\Internet Explorer\iexplore.exe"
End If
   
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "ITC Turborater.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://www.turborater.com/login/"
   lnk.Description = "ITC TurboRater"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_itc_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

If Dir$(d$ & "\" & "JA_PrintForms.lnk") <> "" Then

   'Set lnk = obj.CreateShortcut(d$ & "\" & "JA_PrintForms.lnk")
   
  ' lnk.TargetPath = ruta_internet$
 '  lnk.Arguments = " http://192.168.21.252/index.php/manual-forms/"
   
   
   'lnk.Description = "JA_Print_Forms"
   ' lnk.HotKey = "ALT+CTRL+F"
  ' lnk.IconLocation = ruta_iconos$ + "ja_forms_new.ico"
  ' lnk.WindowStyle = "1"
  ' lnk.WorkingDirectory = "c:\windows"
  ' lnk.Save
  
   'Set lnk = Nothing
    Kill d$ & "\" & "JA_PrintForms.lnk"
End If
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   If Dir$(d$ & "\" & "Ticket System.lnk") <> "" Then
   
       ' Set lnk = obj.CreateShortcut(d$ & "\" & "Ticket System.lnk")
   
       ' lnk.TargetPath = ruta_chrome$
       ' lnk.Arguments = "https://google.com"
            
   
       'lnk.Description = "Unavailable"
       ' lnk.IconLocation = ruta_iconos$ + "ja_not_available.ico"
       ' lnk.WindowStyle = "1"
       ' lnk.WorkingDirectory = "c:\windows"
       ' lnk.Save
   'Clean up
       ' Set lnk = Nothing
        
        Kill d$ & "\" & "Ticket System.lnk"
   End If
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   If Dir$(d$ & "\" & " Zipwhip Texting.lnk") Then
   
   'Set lnk = obj.CreateShortcut(d$ & "\" & " Zipwhip Texting.lnk")
   'lnk.TargetPath = ruta_chrome$
   
   'lnk.Arguments = "https://google.com"
   
  ' lnk.Description = "Unavailable"
   
   'lnk.IconLocation = ruta_iconos$ + "ja_not_available.ico"
   'lnk.WindowStyle = "1"
   'lnk.WorkingDirectory = "c:\windows"
   'lnk.Save
   Kill d$ & "\" & " Zipwhip Texting.lnk"
   End If
   Set lnk = Nothing
   
   
   


Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Clock JA.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://secure5.yourpayrollhr.com/ta/JAI04.clock"
   
   
   'lnk.TargetPath = "https://secure5.yourpayrollhr.com/ta/JAI04.clock"
   'lnk.Arguments = ""
   lnk.Description = "Clock JA"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_time_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing



Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Login JA.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://secure5.yourpayrollhr.com/ta/JAI04.login"
   
   
   'lnk.TargetPath = "https://secure5.yourpayrollhr.com/ta/JAI04.login"
   'lnk.Arguments = ""
   lnk.Description = "Login JA"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "JA_login_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "My Email.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://outlook.office.com"
   
   
   'lnk.TargetPath = "https://mail.justautoins.com/index.php"
   'lnk.Arguments = ""
   lnk.Description = "Email"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_mail2_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
  Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Team.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://teams.microsoft.com"
   
   
   'lnk.TargetPath = "https://mail.justautoins.com/index.php"
   'lnk.Arguments = ""
   lnk.Description = "Teams"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_teams2_new.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
   
   Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Excel.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://www.office.com/launch/excel?auth=2"
   
   
   'lnk.TargetPath = "https://mail.justautoins.com/index.php"
   'lnk.Arguments = ""
   lnk.Description = "Excel"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_excel.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
    Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Word.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://www.office.com/launch/word?auth=2"
   
   
   'lnk.TargetPath = "https://mail.justautoins.com/index.php"
   'lnk.Arguments = ""
   lnk.Description = "Word"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_word.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
     Set obj = CreateObject("WScript.Shell")
d$ = GetSpecialfolder(CSIDL_DESKTOP)

   Set lnk = obj.CreateShortcut(d$ & "\" & "Onedrive.lnk")
   
   lnk.TargetPath = ruta_internet$
   lnk.Arguments = " https://justautoins0-my.sharepoint.com/"
   
   
   'lnk.TargetPath = "https://mail.justautoins.com/index.php"
   'lnk.Arguments = ""
   lnk.Description = "Onedrive"
   ' lnk.HotKey = "ALT+CTRL+F"
   lnk.IconLocation = ruta_iconos$ + "ja_onedrive.ico"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "c:\windows"
   lnk.Save
   'Clean up
   Set lnk = Nothing
   
   
   
   
   'Nuevo objeto de tipo wscript.Shell
Set obj = CreateObject("wscript.Shell")
  
d$ = GetSpecialfolder(CSIDL_DESKTOP)
  
  
'Crea el acceso directo en la carpeta erspecial indicada
Set acceso_directo = obj.CreateShortcut(d$ & "\" & "Transfer_.lnk")  ' .ink
  
  
With acceso_directo
      
    ' Ruta del archivo al cual hacer el acceso directo
    .TargetPath = "C:\Windows\explorer.exe "
    .Arguments = Chr$(34) + "C:\transfer"
    
    
    .WindowStyle = 1
    .WorkingDirectory = "c:\windows"
      
    .IconLocation = "c:\windows\System32\imageres.dll,175"
    
    'Graba el cambio
    .Save
      
End With
  
' Elimina el obeto
Set obj = Nothing
End Sub


Public Sub Checa_status()
On Error Resume Next




End Sub
Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
  contraseña_ini$ = "PAieMu2DLBA6uNj86rSnCDpP"    '"admin"
 user_ini$ = "callc"   '"sa"
 bd_ini$ = "callcenter"   ' "CallCenter"
 server_ini$ = "justautocallcenter.couaea5kjoa1.us-west-1.rds.amazonaws.com"    '"justauto.couaea5kjoa1.us-west-1.rds.amazonaws.com"  '"192.168.21.250"
 
              
 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub



Private Sub btncancel_Click()
On Error Resume Next
base.Close

End
End Sub

Private Sub btncarga_Click()
On Error Resume Next

 b$ = GetIPAddress()
 num_oficina = Mid$(b$, 9, 2)
 
 
 usuario_oficina$ = ""
 usuario_pass$ = ""
 
 Select Case num_oficina
 Case 43  ' compton
   usuario_oficina$ = "Comp"
   usuario_pass$ = "c0mp!"
   
 Case 45  ' covina
   usuario_oficina$ = "Covi"
   usuario_pass$ = "c0v1!"
      
 Case 54  ' echo park
   usuario_oficina$ = "Echo"
   usuario_pass$ = "3ch0!"
   
 Case 49  '  florence
   usuario_oficina$ = "Flor"
   usuario_pass$ = "fl0r!"
   
 Case 84  ' haven
   usuario_oficina$ = "Haven"
   usuario_pass$ = "h4v3n!"
   
 Case 47  ' san bernardino
   usuario_oficina$ = "Bern"
   usuario_pass$ = "b3rn!"
   
 Case 41  ' Santa Ana
   usuario_oficina$ = "Sant"
   usuario_pass$ = "s4nt!"
   
 Case 46  ' Ponderosa
   usuario_oficina$ = "Pond"
   usuario_pass$ = "p0nd!"
   
 Case 23  ' Whittier
   usuario_oficina$ = "Whit"
   usuario_pass$ = "wh1t!"
   
 Case 39
   usuario_oficina$ = "Arle"
   usuario_pass$ = "4rl3!"
   
 End Select
 
 
 txtUser.Text = usuario_oficina$
 txtpass.Text = usuario_pass$
 
 txtUser.SetFocus
 
 
 
End Sub

Public Function GetIPAddress() As String
On Error Resume Next
   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & STR$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function



Private Sub btnclean_Click()
 
 
 txtUser.Text = ""
 txtpass.Text = ""
 
 txtUser.SetFocus
End Sub

Private Sub btnicons_Click()
End Sub

Private Sub btnok_Click()
On Error Resume Next



If txtUser.Text = "" Then
  MsgBox "You need to type the user name", 16, "Attention"
  Exit Sub

End If

If txtpass.Text = "" Then
  MsgBox "You need to type the password", 16, "Attention"
  Exit Sub

End If


If UCase(txtUser.Text) = "ADMIN" Then
   If txtpass.Text = "Tech789" Then
      Hide
      ' crear_iconos
      base.Close
      agente = UCase(txtUser.Text)
      administrador$ = "Y"
      Load forma_main
      forma_main.Show
 
      Unload Me
      Exit Sub
   
   Else
     MsgBox "Password is not valid", 16, "Access denied"
     Exit Sub
   End If

End If


' ************************************************************
  ' carga el campo de usuario


    Set Rs = New ADODB.Recordset
    Checa_status
   
    sSelect = "SELECT login From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    user1$ = Rs(0)
            
    If UCase(user1$) <> UCase(txtUser.Text) Then
      MsgBox "Username does not exist", 16, "Access denied"
      Rs.Close
      Exit Sub
    End If
                         
                         
    Rs.Close
    


 ' ************************************************************
  ' carga el campo de password


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT contrasena From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    password$ = Rs(0)
            
                         
    Rs.Close
    
  
  ' ************************************************************
  ' carga el campo de activo


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT activo From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    activo1$ = Rs(0)
            
                         
    Rs.Close
  
  
 
 If activo1$ = "N" Then
   MsgBox "This user is not active. Please, contact your administrator.", 16, "Access denied"
   Exit Sub
 End If
 
 
 
 
 
 ' ************************************************************
  ' carga el campo de callcenter


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT callcenter From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    callcenter1$ = Rs(0)
            
                         
    Rs.Close
  
  
  
  ' ************************************************************
  ' carga el campo de admin


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT admin From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    administrador$ = Rs(0)
            
                         
    Rs.Close
  
 
 ' ************************************************************
  ' carga el campo de region


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT region From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    region_zona = Rs(0)
            
                         
    Rs.Close
 
 
 
  ' ************************************************************
  ' carga el campo de manager regional


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT tipo_manager From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    regional = Val(Rs(0))
            
                         
    Rs.Close
 
 
 ' ************************************************************
  ' carga el campo de manager


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT manager From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    gerente$ = Rs(0)
            
                         
    Rs.Close
 
 
  ' ************************************************************
  ' carga el campo de full access


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT AccesoTotal From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    full_access1$ = Rs(0)
            
    If full_access1$ = "Y" Then
      regional = 1
      region_zona = 1
    End If
                         
    Rs.Close
 
 
 
 
 

' ************************************************************
  ' carga el campo de abreviatura oficina


    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT nombre_oficina From employees where login='" + UCase(txtUser.Text) + "'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    oficina_autorizada$ = Rs(0)
            
                         
    Rs.Close




If (txtpass.Text) = password$ Then
  Hide
   ' crear_iconos
    'asignar_fondo
    base.Close
    agente = UCase(txtUser.Text)
    
    If callcenter1$ = "Y" Then
   
      Load forma_main
      forma_main.Show
 
    Else
      Load forma_vendor
      forma_vendor.Show
    End If
  
  
  Unload Me
Else
  MsgBox "Password is not valid", 16, "Access denied"
  Exit Sub
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir "c:\callcenter"
MkDir "c:\transfer"

'Conecta_SQL

Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 2000

If (App.PrevInstance = True) Then
  'base.Close
  End
End If


'If Dir$("f:\f\tech\callcenter\version.txt") <> "" Then
  ' base.Close
  
  actualiza = 0
  nf = FreeFile
  Open "\\192.168.84.215\callcenter\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\callcenter\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     R$ = Shell("\\192.168.84.215\callcenter\actualizador.exe", vbNormalFocus)
     
     Hide
     Refresh
     End
     
  End If
  
  
  
'End If



ruta$ = "c:\callcenter\"


      fuente$ = "c:\callcenter\"
      
      transfiere$ = "777"
      Load FrmConfig
      FrmConfig.Show 1
    
       transfiere$ = ""
'      FileCopy App.Path & "\config.ini", fuente$ + "config.ini"
      
Conecta_SQL
      
seg_end = 0
Tipo = 2

Checa_status
End Sub

Private Sub Image3_Click()
ruta$ = "c:\callcenter\"
Image1.Visible = True
End Sub

Private Sub Image4_DblClick()
On Error Resume Next
n$ = App.Path & "\azirotua.dll"


  nf = FreeFile
  Open n$ For Output Shared As #nf
  Lock #nf
  Print #nf, "312443hjdjklwfhwefkljdsa789237hkdk jj sd d 3223at756H23231aasd%21" + LTrim(STR(Val(Format(Now, "yyyy") + 1) * 2))
  Unlock #nf
  Close #nf
  


End Sub


Private Sub Timer1_Timer()
On Error Resume Next
seg_end = seg_end + 1
If seg_end >= 10 Then
  If Dir$(ruta$ + "Abort") <> "" Then
     End
  End If
  seg_end = 0
  
  
End If

If seg_end >= 3 Then
  If actualiza = 1 Then End
End If

End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
  btnok_Click
End If

End Sub



Public Sub asignar_fondo()
On Error Resume Next

  Dim strUserName As String
  
  Dim filename As String
  Dim X As Long
    
    
  strUserName = String(100, Chr$(0))
  'Get the username
  GetUserName strUserName, 100
  'strip the rest of the buffer
  strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
  
  user1$ = strUserName
  
  
  
  
  FileCopy "\\192.168.21.250\fondo\windows.jpg", "c:\transfer\windows.jpg"
  
  c$ = "C:\Users\" + user1$ + "\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.jpg"
  
  Kill c$
  FileCopy "c:\transfer\windows.jpg", c$
  'Kill "c:\transfer\windows.jpg"
  
  Clipboard.Clear
  Clipboard.SetText Left(c$, Len(c$) - 24)
  
  a$ = "copy f:\fondo\windows.jpg " + c$
  
  Call OReg.EstablecerValor(HKEY_CURRENT_USER, "control Panel\Desktop", "wallpaper", c$, REG_SZ)  '
  
  ' asigna la imagen al escritorio
  n$ = "REG add " + Chr$(34) + "HKCU\control Panel\Desktop" + Chr$(34) + " /v wallpaper /t REG_SZ /d " + c$ + " /f"
  ' impide se modifique
  n2$ = "REG add " + Chr$(34) + "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop" + Chr$(34) + " /v NoChangingWallpaper /t REG_DWORD /d " + valor1$ + " /f"
  ' cambia la posicion 0=center, 1=Tile 2=stretch  3=fit  4=fill
  n3$ = "REG add " + Chr$(34) + "HKCU\control Panel\Desktop" + Chr$(34) + " /v WallpaperStyle /t REG_SZ /d 2 /f"


    X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", _
       SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

   filename = c$

    X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, filename, _
       SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
  
End Sub
