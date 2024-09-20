Attribute VB_Name = "Shortcuts"
'Constantes para las carpetas especiales de windows
'----------------------------------------------------
  
'Escritorio de windows
Public Const CSIDL_DESKTOP = &H0
'Carpeta de Inicio - Programas
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_PERSONAL = &H5
'Carpeta Favoritos
Public Const CSIDL_FAVORITES = &H6
'Inicio
Public Const CSIDL_STARTUP = &H7
'Documentos recientes
Public Const CSIDL_RECENT = &H8
  
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_FAVORITES = &H1F
  
  
'Declaraciones del Api
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" ( _
    ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As ITEMIDLIST) As Long
  
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
    ByVal pidl As Long, _
    ByVal pszPath As String) As Long
  
  
Public Const MAX_PATH = 260
  
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
  
  
'Recupera el path de las carpetas y directorios especiales de windows
  
Public Function GetSpecialfolder(CSIDL As Long) As String
Dim Ret As Long, IDL As ITEMIDLIST
Ret = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If Ret = NOERROR Then
       Path$ = Space$(512)
       Ret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
       GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
  
GetSpecialfolder = ""
  
End Function

