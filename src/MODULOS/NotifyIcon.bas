Attribute VB_Name = "NotifyIcon"
 Option Explicit
 'Estructura NOTIFYICONDATA para usar con Shell_NotifyIcon
 Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
 End Type
   
 'Variable para la estructura anterior
 Public sysTray As NOTIFYICONDATA
   
   
 'Constantes
 Public Const NOTIFYICON_VERSION = 3
 Public Const NOTIFYICON_OLDVERSION = 0
   
 Public Const NIM_ADD = &H0
 Public Const NIM_MODIFY = &H1
 Public Const NIM_DELETE = &H2
   
 Public Const NIM_SETFOCUS = &H3
 Public Const NIM_SETVERSION = &H4
   
 Public Const NIF_MESSAGE = &H1
 Public Const NIF_ICON = &H2
 Public Const NIF_TIP = &H4
   
 Public Const NIF_STATE = &H8
 Public Const NIF_INFO = &H10
   
 Public Const NIS_HIDDEN = &H1
 Public Const NIS_SHAREDICON = &H2
   
 Public Const NIIF_NONE = &H0
 Public Const NIIF_WARNING = &H2
 Public Const NIIF_ERROR = &H3
 Public Const NIIF_INFO = &H1
 Public Const NIIF_GUID = &H4
   
 Public Const WM_MOUSEMOVE = &H200
 Public Const WM_LBUTTONDOWN = &H201
 Public Const WM_LBUTTONUP = &H202
 Public Const WM_LBUTTONDBLCLK = &H203
 Public Const WM_RBUTTONDOWN = &H204
 Public Const WM_RBUTTONUP = &H205
 Public Const WM_RBUTTONDBLCLK = &H206
 
'Api SetForegroundWindow Para traer la ventana al frente
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    
' DeclaraciÃ³n Api
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Sub mostrar_systray_notificacion(titulo As String, mensaje As String)
 
 With sysTray
         .cbSize = Len(sysTray)
         .hwnd = WinSeek.Picture_icon.hwnd
         .uID = vbNull
         .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = WinSeek.Picture_icon
         .szTip = WinSeek.Caption & vbNullChar
         .dwState = 0
         .dwStateMask = 0
         .szInfo = mensaje & Chr(0)
         .szInfoTitle = titulo & Chr(0)
         .dwInfoFlags = NIIF_INFO
         .uTimeout = 100
    End With
    'Modifica el ícono con la información
    
    Shell_NotifyIcon NIM_MODIFY, sysTray
    
End Sub
