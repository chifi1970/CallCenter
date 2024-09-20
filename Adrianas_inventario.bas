Attribute VB_Name = "Module1"
Global ruta$, usuario$, codigo$, cadena As String, ruta1$, password$, transfiere$, valido1 As Integer, administrador$
Global region_zona As Integer, oficina_autorizada$, regional As Integer, callcenter1$, full_access1$, gerente$
' para registrar

'Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Public Const REG_SZ = 1 ' Unicode nul terminated String
'Public Const REG_DWORD = 4 ' 32-bit number
'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_USERS = &H80000003
'Public Const HKEY_PERFORMANCE_DATA = &H80000004
'Public Const ERROR_SUCCESS = 0&


  

 


' *************************************************************************************


' esto es para obtener el IP

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
' Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
   



Type bd1
numero As Integer
department As String * 80
category As Integer
subcategory As String * 60
brand As String * 40
model As String * 80
serial As String * 30
Mac As String * 30
procesador As String * 60
ip_address As String * 16
Operating_system As String * 100
Computer_name As String * 60
User As String * 40
descripcion As String * 200
location As String * 80
Last_location As String * 80
estado_registro As Integer
Cantidad As Integer
piezas_x_paquete As Integer

domain As String * 50
memory As String * 20
resolution_monitor As String * 40
size_c As String * 15
modelo_c As String * 60
modelo_d As String * 60
map_net1 As String * 60
map_net2 As String * 60
map_net3 As String * 60
map_net4 As String * 60
map_net5 As String * 60
map_net6 As String * 60
map_net7 As String * 60
map_net8 As String * 60
map_net9 As String * 60
map_net10 As String * 60
port_printer As String * 20
extra1 As String * 20
extra2 As String * 20


nombre As String * 40
fecha_alta As String * 10
fecha_baja As String * 10

fecha_registro As String * 10
nombre_que_inicia_ingreso As String * 40


End Type


Global Const tam_reg = 1981
Global reg As bd1


' fix table
' ================================================================

Type bd3
numero As Integer
department As String * 80
category As Integer
subcategory As String * 60
brand As String * 40
model As String * 80
serial As String * 30
Mac As String * 30
ip_address As String * 16
Operating_system As String * 100
Computer_name As String * 60
User As String * 40
descripcion As String * 200
location As String * 80
Last_location As String * 80
estado_registro As Integer
Cantidad As Integer
piezas_x_paquete As Integer

domain As String * 50
memory As String * 20
resolution_monitor As String * 40
size_c As String * 15
modelo_c As String * 60
modelo_d As String * 60
map_net1 As String * 60
map_net2 As String * 60
map_net3 As String * 60
map_net4 As String * 60
map_net5 As String * 60
map_net6 As String * 60
map_net7 As String * 60
map_net8 As String * 60
map_net9 As String * 60
map_net10 As String * 60
port_printer As String * 20
extra1 As String * 20
extra2 As String * 20


nombre As String * 40
fecha_alta As String * 10
fecha_baja As String * 10

fecha_registro As String * 10
nombre_que_inicia_ingreso As String * 40


End Type


Global Const tam_reg2 = 1921  '1016
Global reg2 As bd3

' =======================================================




Type bd2
nombre As String * 50
pass As String * 20
End Type

Global Const tam_pass = 70
Global pass As bd2

Type nombres_comp1
nombre As String * 60
location As String * 80
department As String * 80
status As Integer
End Type

Global Const tam_computer = 222
Global compu As nombres_comp1


Type sub1
 categoria As Integer
 subcategoria As String * 60
End Type

Global Const tam_sub = 62
Global subcatego As sub1


Type stock2
  categoria As String * 15
  brand As String * 40
  model As String * 80
  stock As Integer
End Type

Global Const tam_stock = 137
Global stock As stock2

Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
       On Error Resume Next
       
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For i = 0 To .Count - 1
         If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).Top = .Controls(i).Top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).Top * SFY, _
            .Controls(i).Width * SFX, _
            .Controls(i).Height * SFY
         End If
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
End Sub



Public Function decodifica(cadena As String)
On Error Resume Next

R$ = ""
a$ = ""
For t = 1 To Len(cadena)
  a$ = Chr$(Asc(Mid(cadena, t, 1)) - 15)
  If Asc(a$) = 17 Then a$ = " "
  R$ = R$ + a$
Next t

codigo$ = R$
End Function




Public Function codifica(cadena As String)
On Error Resume Next

R$ = ""
For t = 1 To Len(cadena)
  R$ = R$ + Chr$(Asc(Mid(cadena, t, 1)) + 15)
Next t

codigo$ = R$



End Function




' *********************************

