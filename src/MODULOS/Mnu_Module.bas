Attribute VB_Name = "Mnu_Module"
Public Sub sub_mnu_about_newfolder_newfolder()
TIULO_MSGBOX = "Acerca de los Virus"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
MsgBox "NewFolder: " & LoadResString(16), vbInformation, TIULO_MSGBOX
KillTimer WinSeek.hwnd, 1
End Sub
Public Sub sub_mnu_about_newfolder_win2x()
TIULO_MSGBOX = "Acerca de los Virus"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
MsgBox "Win2x: " & LoadResString(17), vbInformation, TIULO_MSGBOX
KillTimer WinSeek.hwnd, 1
End Sub
Public Sub sub_mnu_about_newfolder_recycled()
TIULO_MSGBOX = "Acerca de los Virus"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
MsgBox "Recycled: " & LoadResString(18), vbInformation, TIULO_MSGBOX
KillTimer WinSeek.hwnd, 1
End Sub
Public Sub sub_mnu_about_newfoldercleaner()

End Sub


Public Sub eliminar_tools()
On Error GoTo NOPASANADAOSEANOELIMINAMOS:
If GetAttr("tools/flashdesinfector.exe") Then
    Kill "tools/flashdesinfector.exe"
    RmDir "tools"
NOPASANADAOSEANOELIMINAMOS:
End If
End Sub

Public Sub sub_mnu_consejosdeusodelaaplicacion()
'Segundo -----------------------------------------------------------------------------
TIULO_MSGBOX = "Consejos"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
MsgBox "Versión del New Folder." & vbNewLine & _
vbNewLine & _
"El instalador del virus New Folder se copia a su memoria USB un segundo después de ser reconocida por windows. Esto pasa porque el sistema operativo donde conecto su memoria USB estaba infectado con el virus New Folder. Este virus tiene apariencia de carpeta (icono) y pesa 262KB, estos datos son de la versión 3.2.0.1 creada el lunes, 18 de junio de 2007, 05:25:18 a.m." & vbNewLine & _
vbNewLine & _
"New Folder Cleaner ha sido probada con la versión 3.2.0.1 y funciona a la perfección. Si usted sabe de otra versión de este virus, por favor comprímala en formato .zip o .rar y envíemela a la dirección de correo electrónico paulomcnally@yahoo.com y así yo modificaré NewFolder Cleaner a dicha versión.", vbInformation, TIULO_MSGBOX
KillTimer WinSeek.hwnd, 1
End Sub

Public Sub iniciando_la_aplicacion()
TIULO_MSGBOX = "Primer paso"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
If MsgBox("¿Desea desinfectar su PC del virus New Folder?" & vbNewLine & _
vbNewLine & _
"Si su PC está infectada no servirá de nada eliminar los rastros del virus porque se volverán a crear!" & vbNewLine & _
vbNewLine & _
"Si selecciona NO se procederá a la ventana de eliminación de los rastros del virus." & vbNewLine & _
"Si selecciona SI se ejecutará Flash Desinfector, éste eliminará el virus Madre y habilitará todas las opciones deshabilitadas y luego se procederá a la ventana de eliminación de los rastros del virus.", vbYesNo + vbQuestion, TIULO_MSGBOX) = vbYes Then

Shell "tools/FlashDesinfector.exe", vbNormalFocus
End If
KillTimer WinSeek.hwnd, 1
End Sub
