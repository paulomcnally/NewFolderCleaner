Attribute VB_Name = "Apariencia"
'MOVER FORMULARIO
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'APARIENCIA WINDOWS XP
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
'TRANSPARENCIA
 Option Explicit
 ' para crear la transparencia en el MsgBox
 ''''''''''''''''''''''''''''''''''''''''''''
 Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

 ' finaliza el timer creado con SetTimer
 Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

 ' Obtiene el Hwnd del Msgbox
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

 ' Para crear el timer
 Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

 ' para ver si el api SetLayeredWindowAttributes est� presente en el sistema
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'Para deshabilitar el men� y otros
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Obtiene el Handle al men� del sistema de la ventana
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Const MF_BYPOSITION = &H400&
 
 ' constantes
 '''''''''''''''''''''''''''''''
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000


Public hMessageBox As Long 'handle
Public TIULO_MSGBOX As String 'caption



 ' Funci�n para el temporizador
 '''''''''''''''''''''''''''''''''''''''''''''''''''
 Public Sub TimerProc(ByVal hwnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal idEvent As Long, _
                      ByVal dwTime As Long)


         Select Case idEvent
             Case 1
                 ' handle del cuadro de mensaje
                 hMessageBox = FindWindow("#32770", TIULO_MSGBOX)
                 ' comprueba que la funci�n _
                 SetLayeredWindowAttributes se encuentra en windows
If ApiExiste("SetLayeredWindowAttributes", "User32") Then
                     ' aplica la transparencia a este msgbox con un valor de 125
                     aplicar_transparencia hMessageBox, 240
End If
                 ' mata el timer
                 KillTimer hwnd, 1
         End Select
 End Sub

 ' sub para hacer la ventana transparente
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Public Sub aplicar_transparencia(Handle As Long, Valor As Byte)
     Dim Ret As Long

     Ret = GetWindowLong(Handle, GWL_EXSTYLE)
     Ret = Ret Or WS_EX_LAYERED
     SetWindowLong Handle, GWL_EXSTYLE, Ret
     SetLayeredWindowAttributes Handle, 0, Valor, LWA_ALPHA
 End Sub

 Public Function ApiExiste(ByVal NombreFuncion As String, _
                            ByVal NombreDLL As String) As Boolean

     Dim Handle As Long
     Dim Direccion  As Long

     Handle = LoadLibrary(NombreDLL)
If Handle <> 0 Then
         ' si retorna 0 no existe
         Direccion = GetProcAddress(Handle, NombreFuncion)
         FreeLibrary Handle
End If
     ApiExiste = (Direccion <> 0)
 End Function

Sub Main()
    InitCommonControls
    loading.Show
End Sub

'PARA LOS MENU POP UP
Sub Show_Menu_PopUp(El_menu As Object, Button As Integer)

If Button = vbRightButton Then

         Dim El_Form As Form

         ' Referencia al formulario para poder _
          utilizar el m�todo PopupMenu
        Set El_Form = El_menu.Parent

         'Libera el mouse para que no se despliegue el men� est�ndar
         ReleaseCapture

        ' Despliega el men� propio
         El_Form.PopupMenu El_menu

         'Elimina la referencia al formulario
         Set El_Form = Nothing

End If

End Sub

'PARA NO REDIMENCIONAR LA VENTANA
Public Sub Propiedades_del_formulario(ByVal El_Formulario As Form, ByVal Menu_Cerrar As Boolean, ByVal Redimensionar As Boolean, ByVal Mover As Boolean, ByVal Size_Height As Long, ByVal Size_Width As Long)
Dim Hwnd_Menu As Long

'Obtiene el Hwnd del men� para usar con el Api DeleteMenu
Hwnd_Menu = GetSystemMenu(El_Formulario.hwnd, False)

'Tama�o de la ventana
El_Formulario.Height = Size_Height
El_Formulario.Width = Size_Width



'bot�n Cerrar
If Menu_Cerrar Then
   Call DeleteMenu(Hwnd_Menu, 4, MF_BYPOSITION)
End If

'Hace que la ventana no se pueda cambiar de tama�o
If Redimensionar Then
   Call DeleteMenu(Hwnd_Menu, 2, MF_BYPOSITION)
End If

'No permite que la ventana se pueda mover
If Mover Then
   Call DeleteMenu(Hwnd_Menu, 1, MF_BYPOSITION)
End If
End Sub

