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
MsgBox "Versi�n del New Folder." & vbNewLine & _
vbNewLine & _
"El instalador del virus New Folder se copia a su memoria USB un segundo despu�s de ser reconocida por windows. Esto pasa porque el sistema operativo donde conecto su memoria USB estaba infectado con el virus New Folder. Este virus tiene apariencia de carpeta (icono) y pesa 262KB, estos datos son de la versi�n 3.2.0.1 creada el lunes, 18 de junio de 2007, 05:25:18 a.m." & vbNewLine & _
vbNewLine & _
"New Folder Cleaner ha sido probada con la versi�n 3.2.0.1 y funciona a la perfecci�n. Si usted sabe de otra versi�n de este virus, por favor compr�mala en formato .zip o .rar y env�emela a la direcci�n de correo electr�nico paulomcnally@yahoo.com y as� yo modificar� NewFolder Cleaner a dicha versi�n.", vbInformation, TIULO_MSGBOX
KillTimer WinSeek.hwnd, 1
End Sub

Public Sub iniciando_la_aplicacion()
TIULO_MSGBOX = "Primer paso"
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
If MsgBox("�Desea desinfectar su PC del virus New Folder?" & vbNewLine & _
vbNewLine & _
"Si su PC est� infectada no servir� de nada eliminar los rastros del virus porque se volver�n a crear!" & vbNewLine & _
vbNewLine & _
"Si selecciona NO se proceder� a la ventana de eliminaci�n de los rastros del virus." & vbNewLine & _
"Si selecciona SI se ejecutar� Flash Desinfector, �ste eliminar� el virus Madre y habilitar� todas las opciones deshabilitadas y luego se proceder� a la ventana de eliminaci�n de los rastros del virus.", vbYesNo + vbQuestion, TIULO_MSGBOX) = vbYes Then

Shell "tools/FlashDesinfector.exe", vbNormalFocus
End If
KillTimer WinSeek.hwnd, 1
End Sub
