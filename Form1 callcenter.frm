VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00F8E4D8&
   Caption         =   "by Leandro Ascierto"
   ClientHeight    =   10740
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   3855
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   23
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k1"
            Object.ToolTipText     =   "Negrita"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k2"
            Object.ToolTipText     =   "Cursiva"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k3"
            Object.ToolTipText     =   "Subrayado"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k4"
            Object.ToolTipText     =   "Color de Fuente"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k5"
            Object.ToolTipText     =   "Color de Fondo"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k6"
            Object.ToolTipText     =   "Insertar Imágen"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k7"
            Object.ToolTipText     =   "Insertar Hipervínculo"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k8"
            Object.ToolTipText     =   "Insertar Línea"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k9"
            Object.ToolTipText     =   "Insertar Tabla"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k10"
            Object.ToolTipText     =   "Alinear a la Izquierda"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k11"
            Object.ToolTipText     =   "Centrar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k12"
            Object.ToolTipText     =   "Insertar a la Derecha"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k13"
            Object.ToolTipText     =   "Justificar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k14"
            Object.ToolTipText     =   "Aumentar Sangría"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k15"
            Object.ToolTipText     =   "Disminuir Sangría"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k16"
            Object.ToolTipText     =   "Numeración"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k17"
            Object.ToolTipText     =   "Viñetas"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k18"
            Object.ToolTipText     =   "Iconos Gestuales"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin Project1.CoolComboBox CoolComboBox2 
         Height          =   255
         Left            =   9120
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.CoolComboBox CoolComboBox1 
         Height          =   255
         Left            =   7200
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
   End
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1164
      ButtonWidth     =   1693
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Enviar       "
            Key             =   "k119"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Adjuntar     "
            Key             =   "k120"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Imprimir    "
            Key             =   "k121"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Configurar"
            Key             =   "k122"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Prioridad "
            Key             =   "k123"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicConteiner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   4
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   12615
      TabIndex        =   8
      Top             =   2640
      Width           =   12615
      Begin ComctlLib.ListView ListView1 
         Height          =   975
         Left            =   1080
         TabIndex        =   14
         Top             =   120
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   1720
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjuntos:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.PictureBox PicConteiner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   7
      Top             =   2145
      Width           =   12615
      Begin VB.TextBox TxtAsunto 
         Height          =   380
         Left            =   1080
         TabIndex        =   12
         Top             =   120
         Width           =   10575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox PicConteiner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   6
      Top             =   1650
      Visible         =   0   'False
      Width           =   12615
      Begin VB.TextBox TxtCCO 
         Height          =   380
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   10575
      End
      Begin Project1.ButtonOffice ButtonOffice1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Enabled         =   0   'False
         Caption         =   "Boton1"
         BackColor       =   -2147483633
      End
   End
   Begin VB.PictureBox PicConteiner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   5
      Top             =   1155
      Width           =   12615
      Begin VB.TextBox TxtCC 
         Height          =   380
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   10575
      End
      Begin VB.PictureBox Picture2 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   18
         Top             =   0
         Width           =   0
      End
      Begin Project1.ButtonOffice ButtonOffice1 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Enabled         =   0   'False
         Caption         =   "Boton1"
         BackColor       =   -2147483633
      End
   End
   Begin VB.PictureBox PicConteiner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   4
      Top             =   660
      Width           =   12615
      Begin Project1.ButtonOffice ButtonOffice1 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Enabled         =   0   'False
         Caption         =   "Boton1"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox TxtPara 
         Height          =   380
         Left            =   1080
         TabIndex        =   11
         Top             =   120
         Width           =   10575
      End
   End
   Begin VB.PictureBox ctxHookMenu1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11880
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   4200
      Width           =   1200
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   12615
      TabIndex        =   1
      Top             =   10410
      Width           =   12615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5640
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   9948
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageListSmyles 
      Left            =   11880
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   16777215
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11880
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Configurar Impresora"
         Index           =   2
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Vista previa"
         Index           =   3
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Imprimir"
         Index           =   4
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Salir"
         Index           =   6
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Begin VB.Menu SubMnuVer 
         Caption         =   "Campo CC"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu SubMnuVer 
         Caption         =   "Campo CCO"
         Index           =   1
      End
   End
   Begin VB.Menu PopUpAdjunto 
      Caption         =   "PopUpAdjunto"
      Visible         =   0   'False
      Begin VB.Menu SubMnuAdjuntar 
         Caption         =   "Abrir"
         Index           =   0
      End
      Begin VB.Menu SubMnuAdjuntar 
         Caption         =   "Adjuntar"
         Index           =   1
      End
      Begin VB.Menu SubMnuAdjuntar 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu SubMnuAdjuntar 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubMnuAdjuntar 
         Caption         =   "Propiedades"
         Index           =   4
      End
   End
   Begin VB.Menu PopUpEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Cortar"
         Index           =   0
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Copiar"
         Index           =   1
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Pegar"
         Index           =   2
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Eliminar"
         Index           =   3
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Seleccionar Todo"
         Index           =   5
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Deshacer"
         Index           =   7
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Rehacer"
         Index           =   8
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Buscar"
         Index           =   10
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu SubMnuBrowser 
         Caption         =   "Propiedades"
         Index           =   12
      End
   End
   Begin VB.Menu MnuFormato 
      Caption         =   "&Formato"
      Begin VB.Menu SubMnuFormato 
         Caption         =   "Fondo"
         Begin VB.Menu SubMnuFondo 
            Caption         =   "Imagen"
            Index           =   0
         End
         Begin VB.Menu SubMnuFondo 
            Caption         =   "Color"
            Index           =   1
         End
      End
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu SubMnuOpciones 
         Caption         =   "Skin"
         Index           =   0
         Begin VB.Menu SubMnuSkin 
            Caption         =   "Skin Azul"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu SubMnuSkin 
            Caption         =   "Skin Gris"
            Index           =   1
         End
      End
      Begin VB.Menu SubMnuOpciones 
         Caption         =   "Solicitar confirmación de lectura"
         Index           =   1
      End
   End
   Begin VB.Menu MnuPrioridad 
      Caption         =   "MnuPrioridad"
      Visible         =   0   'False
      Begin VB.Menu SubMnuPrioridad 
         Caption         =   "Baja"
         Index           =   0
      End
      Begin VB.Menu SubMnuPrioridad 
         Caption         =   "Normal"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu SubMnuPrioridad 
         Caption         =   "Alta"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long


'Estructura MENUINFO
Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    RhbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

' Para extraer el ícono de los adjuntos
Private Type IconType
  cbSize As Long
  picType As PictureTypeConstants
  hIcon As Long
End Type

Private Type CLSIdType
  id(16) As Byte
End Type

Private Type ShellFileInfoType
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


' para el cuadro de diálogo de win para abrir archivo
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_ENABLETEMPLATE As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_NOLONGNAMES As Long = &H40000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_READONLY As Long = &H1
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH As Long = 2
Private Const OFN_SHAREWARN As Long = 0
Private Const OFN_SHARENOWARN As Long = 1
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFS_MAXPATHNAME As Long = 260

Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS

Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY
             
'apis para las propiedades del adjunto
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const SW_SHOWNORMAL = 1


Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000



Private Enum eSizeIcon
    [Small] = 257
    [Normal] = 256
End Enum

Public WithEvents HTML             As HTMLDocument
Attribute HTML.VB_VarHelpID = -1

Private c_ToolBar1                  As cSubclassToolBar
Private c_ToolBar2                  As cSubclassToolBar
Private mcIni                       As clsIni
Private Element                     As Object
Private PathDocumento               As String
Private mPrioridad                  As Integer



Private Function Establecer_Color_Menu(ByVal hwndfrm As Long, ByVal Color As Long, ByVal subMenu As Boolean) As Boolean

    Dim mi As MENUINFO
    Dim flags As Long

    flags = MIM_BACKGROUND

    If subMenu Then
        flags = flags Or MIM_APPLYTOSUBMENUS
    End If

    With mi
        .cbSize = Len(mi)
        .fMask = flags
        .RhbrBack = CreateSolidBrush(Color)
    End With

    Call SetMenuInfo(GetMenu(hwndfrm), mi)
    Call DrawMenuBar(hwndfrm)

End Function


Private Function ShowProps(filename As String) As Boolean
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .lpVerb = "properties"
        .lpFile = filename
    End With
    ShowProps = ShellExecuteEx(SEI)
End Function

Public Function GetStringFile(value As String, Optional Force As Boolean = True) As String
    Dim HEX As String, Ret As Long
    
    If Force Then value = Replace(value, "+", " ")
    value = Replace(value, "%25", Chr(0))
    
    Ret = InStr(value, "%")
    
    Do While Ret > 0
        Ret = InStr(value, "%")
        If Ret <> 0 Then
            HEX = Mid(value, Ret + 1, 2)
            value = Replace(value, "%" & HEX, Chr("&H" & HEX))
        End If
    Loop
    
    value = Replace(value, Chr(0), "%")
    
    value = Replace(value, "/", "\")
    
    GetStringFile = Replace(value, " ", "_")
End Function


Private Function GetFileNameURL(URL As String)
Dim Ret As Long
Dim sName As String
Ret = InStrRev(URL, "/")
If Ret Then
sName = Mid(URL, Ret + 1)
GetFileNameURL = GetStringFile(sName)
End If
End Function


' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Función para devolver en un IPictureDisp la imagen del ícono de los archivos adjuntos y para el Mp3
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIconFile(sPathFile As String, lSize As eSizeIcon) As IPictureDisp
  
  Dim Ret As Long
  Dim Unkown As IUnknown
  Dim Icon As IconType
  Dim CLSID As CLSIdType
  Dim ShellInfo As ShellFileInfoType
  
  
  'lSize = Small  ' 256 para el ícono chico
  'lSize = Normal ' 257 para el ícono grande
  
  Call SHGetFileInfo(sPathFile, 0, ShellInfo, Len(ShellInfo), lSize)
  
  With Icon
    .cbSize = Len(Icon)
    .picType = vbPicTypeIcon
    .hIcon = ShellInfo.hIcon
  End With
  
  With CLSID
    .id(8) = &HC0
    .id(15) = &H46
  End With
  
  Ret = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
  Set GetIconFile = Unkown ' retornar la imagen
    
End Function

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recuperar solo el nombre de archivo de el path ( para los arhivos adjuntos en el control listview )
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFileName(sPath As String, sChar As String) As String
    Dim Ret As String
    GetFileName = Right(sPath, Len(sPath) - InStrRev(sPath, sChar))
End Function



' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Agregar ícono del archivo al imagelist y el archivo al Listview
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFile(strPath As String, lv As ListView)
    
    On Error GoTo error_handler
    
    ' Ver si existe la imagen .. si no existe agregarla
    If ImageList1.ListImages(strPath) Is Nothing Then
       ImageList1.ListImages.add , strPath, GetIconFile(strPath, Small)
       With lv
           ' asignar el ImageList al LV
           If .SmallIcons Is Nothing Then
              .SmallIcons = ImageList1
              .Icons = ImageList1
           End If
       End With
       
    End If
    
    Dim lvItem As ListItem
    
    If lv.ListItems(strPath) Is Nothing Then
        ' Añadir el nombre del archivo y en la clave la ruta
        Set lvItem = lv.ListItems.add(, strPath, GetFileName(strPath, "\") & " (" & GetFormatKB(FileLen(strPath)) & ")", strPath, strPath)
    End If
    
    Exit Sub
error_handler:
    If Err.Number = 35601 Then ' El item no existe, seguir y agregarlo
       Resume Next
    End If
End Sub


Public Function GetFormatKB(Bytes As Variant) As String
If Bytes >= 0 Then GetFormatKB = Format((Bytes \ 1024) + 1, "##,###,##0") & " KB"
End Function




Private Sub cdmail_Error(descripcion As String, numero As Variant)
If numero Then
    MsgBox "Error " & numero & vbCrLf & vbCrLf & descripcion
Else
    MsgBox descripcion
End If
End Sub



Private Function Decimal_Hex(ColorDecimal As Long) As String
On Error Resume Next
Dim R As String, G As String, b As String

R = CStr(HEX(ColorDecimal And 255))
G = CStr(HEX((ColorDecimal And 65280) / 256))
b = CStr(HEX((ColorDecimal And 16711680) / 65536))

If Len(R) = 1 Then R = R & "0"
If Len(G) = 1 Then G = G & "0"
If Len(b) = 1 Then b = b & "0"

Decimal_Hex = "#" & R & G & b
    
End Function







Private Sub CoolComboBox1_DropDown()
Dim Ret As String
    CoolComboBox1.State = CB_Presed
    Ret = ShowMenuFontList(Me.hWnd, CoolComboBox1.Left, CoolComboBox1.Top + CoolComboBox1.Height + Toolbar1.Top + 30)
    If Ret <> "" Then
        CoolComboBox1.Text = Ret
        HTML.execCommand "FontName", True, Ret
        WebBrowser1.SetFocus
    End If

    CoolComboBox1.State = CB_Normal
End Sub

Private Sub CoolComboBox2_DropDown()
Dim Ret As Integer
    CoolComboBox2.State = CB_Presed
    Ret = ShowMenuFontSize(Me.hWnd, CoolComboBox2.Left, CoolComboBox2.Top + CoolComboBox2.Height + Toolbar1.Top + 30)
    Debug.Print Ret
    If Ret <> 0 Then
        CoolComboBox2.Text = Ret
        HTML.execCommand "FontSize", True, Ret
        WebBrowser1.SetFocus
    End If
    CoolComboBox2.State = CB_Normal
End Sub



Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim DirSmyles As String
Dim File As String

Set mcIni = New clsIni

Set c_ToolBar1 = New cSubclassToolBar
Set c_ToolBar2 = New cSubclassToolBar



For i = 101 To 118
    ImageList1.ListImages.add , "k" & i - 100, LoadResPicture(i, vbResIcon)
Next

Toolbar1.Imagelist = ImageList1

For i = 1 To Toolbar1.Buttons.Count
    If Toolbar1.Buttons(i).Key <> "" Then
        Toolbar1.Buttons(i).Image = Toolbar1.Buttons(i).Key
    End If
Next


For i = 119 To 123
    ImageList2.ListImages.add , "k" & i, LoadResPicture(i, vbResIcon)
Next

Toolbar2.Imagelist = ImageList2

For i = 1 To 6
    If Toolbar2.Buttons(i).Key <> "" Then
        Toolbar2.Buttons(i).Image = Toolbar2.Buttons(i).Key
    End If
Next

DirSmyles = Dir(App.Path & "\Smyles\")

Do While DirSmyles <> ""
    File = App.Path & "\Smyles\" & DirSmyles
    ImageListSmyles.ListImages.add , File, LoadPicture(File)
    DirSmyles = Dir
Loop

c_ToolBar1.SubClassToolBar Toolbar1.hWnd
c_ToolBar2.SubClassToolBar Toolbar2.hWnd, True

SubMnuSkin_Click 0
WebBrowser1.RegisterAsDropTarget = False

If App.LogMode = 0 Then
    WebBrowser1.Navigate App.Path & "/Nueva.htm"
Else
    WebBrowser1.Navigate "res://" & App.Path & "\" & App.EXEName & ".exe/HTML_0"
End If

Do: DoEvents: Loop Until WebBrowser1.readyState = READYSTATE_COMPLETE
WebBrowser1.Document.designMode = "On"
Set HTML = WebBrowser1.Document


CoolComboBox1.Text = "Times New Roman"
CoolComboBox2.Text = "3"
mPrioridad = 1
End Sub

Private Sub AddTable(ByVal X As Long, Y As Long)
    Dim i As Integer, j As Integer
    Dim CodHtml As String
    Dim Range As Object
    
    CodHtml = "<TABLE cellSpacing=1 cellPadding=1 width='100%' border=1>"
    For i = 1 To Y
        CodHtml = CodHtml & "<tr>"
            For j = 1 To X
                CodHtml = CodHtml & "<td width='" & 100 / X & "%'>&nbsp;</td>"
            Next
        CodHtml = CodHtml & "</tr>"
    Next
    CodHtml = CodHtml & "</table>"
    
    Set Range = HTML.selection.createRange
    Range.pasteHTML CodHtml
    
    Set Range = Nothing
End Sub


Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    WebBrowser1.Move -30, Toolbar1.Top + Toolbar1.Height - 30, Me.ScaleWidth + 60, Me.ScaleHeight - Toolbar1.Top - Toolbar1.Height - Picture1.Height + 60
    TxtPara.Width = Me.ScaleWidth - 1200
    TxtCC.Width = Me.ScaleWidth - 1200
    TxtCCO.Width = Me.ScaleWidth - 1200
    TxtAsunto.Width = Me.ScaleWidth - 1200
    ListView1.Width = Me.ScaleWidth - 1200
End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'If PathDocumento <> "" Then DeleteFolder PathDocumento
    Set c_ToolBar1 = Nothing
    Set c_ToolBar2 = Nothing
End Sub



Private Function HTML_oncontextmenu() As Boolean
    PopupMenu PopUpEdicion
End Function

Private Sub PopUpEdicion_Click()
On Error Resume Next
    SubMnuBrowser(0).Enabled = HTML.queryCommandEnabled("Cut")
    SubMnuBrowser(1).Enabled = HTML.queryCommandEnabled("Copy")
    SubMnuBrowser(2).Enabled = HTML.queryCommandEnabled("Paste")
    SubMnuBrowser(3).Enabled = HTML.queryCommandEnabled("Delete")
    SubMnuBrowser(5).Enabled = HTML.queryCommandEnabled("SelectAll")
    SubMnuBrowser(7).Enabled = HTML.queryCommandEnabled("Undo")
    SubMnuBrowser(8).Enabled = HTML.queryCommandEnabled("Redo")
    SubMnuBrowser(10).Enabled = Len(HTML.body.innerText)
    Set Element = HTML.parentWindow.event.srcElement
    If Not Element Is Nothing Then
        If Element.tagName = "IMG" Or Element.tagName = "A" Then
            SubMnuBrowser(12).Enabled = True
        Else
            SubMnuBrowser(12).Enabled = False
        End If
    Else
        SubMnuBrowser(12).Enabled = False
    End If
End Sub


Private Sub HTML_onkeydown()
    Call EstadoBotones
End Sub

Private Function HTML_onclick() As Boolean
    Call EstadoBotones
    HTML_onclick = True
End Function

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim InItem As Boolean
If Button = 2 Then
    InItem = Not ListView1.HitTest(X, Y) Is Nothing
    If InItem Then
        Set ListView1.SelectedItem = ListView1.HitTest(X, Y)
    End If
    SubMnuAdjuntar(0).Enabled = InItem
    SubMnuAdjuntar(2).Enabled = InItem
    SubMnuAdjuntar(4).Enabled = InItem
    PopupMenu PopUpAdjunto
End If
End Sub

Private Sub Picture1_Paint()
Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, IIf(SubMnuSkin(0).Checked, 0, 10), 0, 10
End Sub

Private Sub SubMnuAdjuntar_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute Me.hWnd, vbNullString, ListView1.SelectedItem.Key, vbNullString, "C:\", SW_SHOWNORMAL
        Case 1
            Call Adjuntar
        Case 2
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        Case 4
            ShowProps ListView1.SelectedItem.Key
    End Select
End Sub


Private Sub SubMnuArchivo_Click(Index As Integer)
    Select Case Index
        Case 0
            WebBrowser1.Refresh
        Case 2
            WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
        Case 3
            WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
        Case 4
            WebBrowser1.Document.execCommand "Print"
        Case 6
            Unload Me
    End Select
End Sub

Private Sub SubMnuBrowser_Click(Index As Integer)
    Select Case Index
        Case 0
            HTML.execCommand "Cut"
        Case 1
            HTML.execCommand "Copy"
        Case 2
            HTML.execCommand "Paste"
        Case 3
            HTML.execCommand "Delete"
        Case 5
            HTML.execCommand "SelectAll"
        Case 7
            HTML.execCommand "Undo"
        Case 8
            HTML.execCommand "Redo"
        Case 10
            WebBrowser1.SetFocus
            'SendKeys ("^f")
        Case 12
        If Element.tagName = "IMG" Then HTML.execCommand "InsertImage", True
        If Element.tagName = "A" Then HTML.execCommand "CreateLink", True
    End Select
End Sub


Private Sub EstadoBotones()
On Error Resume Next
    With Toolbar1
        .Buttons(1).value = Abs(HTML.queryCommandValue("bold"))
        .Buttons(2).value = Abs(HTML.queryCommandValue("Italic"))
        .Buttons(3).value = Abs(HTML.queryCommandValue("Underline"))
        .Buttons(12).value = Abs(HTML.queryCommandValue("JustifyLeft"))
        .Buttons(13).value = Abs(HTML.queryCommandValue("JustifyCenter"))
        .Buttons(14).value = Abs(HTML.queryCommandValue("JustifyRight"))
        .Buttons(15).value = Abs(HTML.queryCommandValue("JustifyFull"))
        .Buttons(19).value = Abs(HTML.queryCommandValue("InsertOrderedList"))
        .Buttons(20).value = Abs(HTML.queryCommandValue("InsertUnorderedList"))
    End With
    
    CoolComboBox1.Text = HTML.queryCommandValue("FontName")
    CoolComboBox2.Text = HTML.queryCommandValue("FontSize")
End Sub

Private Sub SubMnuFondo_Click(Index As Integer)
    Dim Color As Long
    
    If Index = 0 Then
        FrmBackGround.Show vbModal, Me
    Else
        Color = ShowDialogColor(Me.hWnd)
        If Color <> -1 Then
            HTML.bgColor = Decimal_Hex(Color)
        End If
    End If
End Sub

Private Sub SubMnuSkin_Click(Index As Integer)
Dim i As Integer
    If Index = 1 Then
        SubMnuSkin(0).Checked = False
        'ctxHookMenu1.MenuLook = MenuXP
        CoolComboBox1.Style = Word2000
        CoolComboBox2.Style = Word2000
        Call Establecer_Color_Menu(Me.hWnd, &HD8E9EC, True)
        Me.BackColor = &HD8E9EC   'vbButtonFace
        
        c_ToolBar1.SetSkinPicture LoadResPicture("BITMAP_1", vbResBitmap) 'LoadPicture(App.Path & "\Skin_Word2000.bmp")
        c_ToolBar2.SetSkinPicture LoadResPicture("BITMAP_1", vbResBitmap) 'LoadPicture(App.Path & "\Skin_Word2000.bmp")
        For i = 0 To 4
            PicConteiner(i).BackColor = &HD8E9EC  ' vbButtonFace
        Next
        For i = 0 To 2
            ButtonOffice1(i).BackColor = &HD8E9EC  ' vbButtonFace
            ButtonOffice1(i).Style = BT_2000
        Next
    Else
        SubMnuSkin(1).Checked = False
        'ctxHookMenu1.MenuLook = Menu2003
        CoolComboBox1.Style = VBNet
        CoolComboBox2.Style = VBNet
        Call Establecer_Color_Menu(Me.hWnd, &HF8E4D8, True)
        Me.BackColor = &HF8E4D8
        c_ToolBar1.SetSkinPicture LoadResPicture("BITMAP_0", vbResBitmap) '           LoadPicture(App.Path & "\Skin_VBNET.bmp")
        c_ToolBar2.SetSkinPicture LoadResPicture("BITMAP_0", vbResBitmap)
        For i = 0 To 4
            PicConteiner(i).BackColor = &HF8E4D8
        Next
        For i = 0 To 2
            ButtonOffice1(i).BackColor = &HF8E4D8
            ButtonOffice1(i).Style = BT_2003
        Next
    End If
    SubMnuSkin(Index).Checked = True
    Picture1.Refresh
End Sub

Private Sub SubMnuPrioridad_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        SubMnuPrioridad(i).Checked = Index = i
    Next
    mPrioridad = Index
End Sub
Private Sub SubMnuOpciones_Click(Index As Integer)

If Index = 1 Then
SubMnuOpciones(Index).Checked = Not SubMnuOpciones(Index).Checked
End If
End Sub

Private Sub SubMnuVer_Click(Index As Integer)
SubMnuVer(Index).Checked = Not SubMnuVer(Index).Checked
PicConteiner(1).Visible = SubMnuVer(0).Checked
PicConteiner(2).Visible = SubMnuVer(1).Checked
WebBrowser1.Move 0, Toolbar1.Top + Toolbar1.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar1.Top - Toolbar1.Height - Picture1.Height
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next

    Dim Ret As Long

    Select Case Button.Index
        Case 1
            HTML.execCommand "Bold"
        Case 2
            HTML.execCommand "Italic"
        Case 3
            HTML.execCommand "Underline"
        Case 5
            Button.value = tbrPressed
            Ret = ShowMenuPaleteColor(Me.hWnd, Button.Left, Button.Top + Button.Height + Toolbar1.Top)
            If Ret > -1 Then
                HTML.execCommand "ForeColor", True, Ret
            End If
            Button.value = tbrUnpressed

        Case 6
            Button.value = tbrPressed
            Ret = ShowMenuPaleteColor(Me.hWnd, Button.Left, Button.Top + Button.Height + Toolbar1.Top)
            If Ret > -1 Then
                HTML.execCommand "BackColor", True, Ret
            End If
            Button.value = tbrUnpressed
        Case 7
            HTML.execCommand "InsertImage", True
        Case 8
            HTML.execCommand "CreateLink", True
        Case 9
            HTML.execCommand "InsertHorizontalRule", True
        Case 10
            Button.value = tbrPressed
            If ShowMenuGrid(Me.hWnd, Button.Left, Button.Top + Button.Height + Toolbar1.Top) Then
                AddTable TablaX, TablaY
            End If
            Button.value = tbrUnpressed
        Case 12
            HTML.execCommand "JustifyLeft"
        Case 13
            HTML.execCommand "JustifyCenter"
        Case 14
            HTML.execCommand "JustifyRight"
        Case 15
            HTML.execCommand "JustifyFull"
        Case 17
            HTML.execCommand "Indent"
        Case 18
            HTML.execCommand "Outdent"
        Case 19
            HTML.execCommand "InsertOrderedList"
        Case 20
            HTML.execCommand "InsertUnorderedList"
        Case 22
            Button.value = tbrPressed
            Ret = ShowMenuSmyles(Me.hWnd, Button.Left, Button.Top + Button.Height + Toolbar1.Top, ImageListSmyles)
            If Ret > 0 Then
                 HTML.execCommand "InsertImage", False, ImageListSmyles.ListImages(Ret).Key
            End If
            Button.value = tbrUnpressed
        End Select
        Toolbar1.Refresh
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1
            Call Enviar
        Case 2
            Call Adjuntar
        Case 3
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        Case 5
            FrmConfig.Show vbModal
        Case 6
            Button.value = tbrPressed
            PopupMenu MnuPrioridad, , Button.Left, Button.Top + Button.Height
            Button.value = tbrUnpressed
    End Select
    Toolbar2.Refresh
End Sub


Private Sub Enviar()
On Error GoTo Informar:
Const sch = "http://schemas.microsoft.com/cdo/configuration/"

Dim INI_PATH As String
Dim loCfg As Object
Dim loMsg As Object
Dim loBP As Object
Dim i As Long
Dim DestImg As String
Dim TempHTML As String
Dim TempHTMLMail As String
Dim strImg As String



INI_PATH = App.Path & "\config.ini"


Me.Enabled = False
FrmProgress.Show , Me
DoEvents

Set loCfg = CreateObject("CDO.Configuration")

With loCfg.Fields
  .Item(sch & "smtpserver") = mcIni.getValue(INI_PATH, "datos", "servidor")
  .Item(sch & "smtpserverport") = mcIni.getValue(INI_PATH, "datos", "puerto")
  .Item(sch & "sendusing") = 2
  .Item(sch & "sendusername") = mcIni.getValue(INI_PATH, "datos", "usuario")
  .Item(sch & "sendpassword") = mcIni.Encriptar(App.EXEName, mcIni.getValue(INI_PATH, "datos", "password"), 2)
  .Item(sch & "smtpusessl") = mcIni.getValue(INI_PATH, "datos", "ssl", 0)
  .Item(sch & "smtpconnectiontimeout") = 15
  .Item(sch & "smtpauthenticate") = mcIni.getValue(INI_PATH, "datos", "Aut", 0)
End With

loCfg.Fields.Update

Set loMsg = CreateObject("CDO.Message")

With loMsg
  .Configuration = loCfg
  .From = mcIni.getValue(INI_PATH, "datos", "usuario")
  .To = TxtPara

  .Subject = TxtAsunto
  
  TempHTML = HTML.documentElement.outerHTML
  
  If HTML.All.tags("BODY")(0).background <> "" Then
  
        strImg = HTML.All.tags("BODY")(0).background
        
        If PathIsURL(strImg) <> 1 Then
            HTML.All.tags("BODY")(0).background = "cid:" & "BackGround"
            
            Set loBP = .AddRelatedBodyPart(strImg, "BackGround", 1)
            
            With loBP.Fields
                .Item("urn:schemas:mailheader:Content-ID") = "BackGround"
                .Update
            End With
        End If
        
  End If
  
  
  For i = 0 To HTML.images.Length - 1
  
        strImg = HTML.images.Item(i).src
        
        If Left(strImg, 8) = "file:///" Then
            DestImg = GetFileNameURL(strImg)
            If DestImg <> "" Then
    
                HTML.images.Item(i).src = "cid:" & DestImg & i
    
                Set loBP = .AddRelatedBodyPart(strImg, DestImg & i, 1)
                
                With loBP.Fields
                    .Item("urn:schemas:mailheader:Content-ID") = DestImg & i
                    .Update
                End With
                
            End If
        End If
    Next

    .HTMLBody = HTML.documentElement.outerHTML

    HTML.body.innerHTML = TempHTML
    
    
    For i = 1 To ListView1.ListItems.Count
        .AddAttachment ListView1.ListItems(i).Key
    Next
    
    
    'Prioridad
    ' -1=Low, 0=Normal, 1=High
    .Fields("urn:schemas:httpmail:priority") = mPrioridad - 1
    .Fields("urn:schemas:mailheader:X-Priority") = mPrioridad - 1
    'Importancia
    '0=Low, 1=Normal, 2=High
    .Fields("urn:schemas:httpmail:importance") = mPrioridad
    
    'Solicitar confirmación de lectura
    If SubMnuOpciones(1).Checked Then
        .Fields("urn:schemas:mailheader:disposition-notification-to") = .From
        .Fields("urn:schemas:mailheader:return-receipt-to") = .From
    End If
    
    .Fields.Update
    .Send
End With


Me.Enabled = True
Unload FrmProgress
  
  
Exit Sub
Informar:
    MsgBox Err.Description
    Me.Enabled = True
    Unload FrmProgress
End Sub


Private Sub Adjuntar()

On Error GoTo err_handler
Dim OFN As OPENFILENAME
Dim vFiles() As String
Dim lFile As Long

   With OFN
      .nStructSize = Len(OFN)
      .hwndOwner = Form1.hWnd
      .sFilter = "Todos los archivos" & vbNullChar & "*.*" & vbNullChar & vbNullChar
      .sFile = Space$(1024) & vbNullChar & vbNullChar
      .nMaxFile = Len(.sFile)
      .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .nMaxTitle = Len(OFN.sFileTitle)
      .flags = OFS_FILE_OPEN_FLAGS Or OFN_ALLOWMULTISELECT
   End With
   
   
    If GetOpenFileName(OFN) Then
    
        vFiles = Split(OFN.sFile, Chr(0))
    
        For lFile = 1 To UBound(vFiles) - 4 ' More than 1 file then do this until there are no more files
            Call AddFile(vFiles(0) + "\" & vFiles(lFile), ListView1)
        Next
    
        If UBound(vFiles) = 4 Then
            Call AddFile(OFN.sFile, ListView1)
        End If
    End If

        
Exit Sub
err_handler:
    MsgBox Err.Description, vbCritical, "Error al añadir archivo como adjunto"
End Sub
