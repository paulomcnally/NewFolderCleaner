Attribute VB_Name = "Desinfeccion"
'Api para eliminar archivos
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Sub Comprobando_PC_en_busca_de_Virus()
'=========================================================
' NewFolder 1
'=========================================================
If FileExist(Environ("windir") & "\system32\RVHOST.exe") Then
    Dim NewFolder1 As String
    NewFolder1 = Environ("windir") & "\system32\RVHOST.exe"
    'Matamos al proceso
    Shell "taskkill /f /im RVHOST.exe", vbHide
    'Le quitamos los atributos
    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & NewFolder1 & WinSeek.doble_comillas.Text, vbHide
    'Eliminamos la entrada que crea en Regedit para AutoEjecutarse
    limpiando_llave_de_inicio_del_virus
    mostrar_systray_notificacion LoadResString(33), LoadResString(26)
End If
'=========================================================
' NewFolder 2
'=========================================================
If FileExist(Environ("windir") & "\RVHOST.exe") Then
    Dim NewFolder2 As String
    NewFolder2 = Environ("windir") & "\RVHOST.exe"
    'Matamos al proceso
    Shell "taskkill /f /im RVHOST.exe", vbHide
    'Le quitamos los atributos
    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & NewFolder2 & WinSeek.doble_comillas.Text, vbHide
    'Eliminamos la entrada que crea en Regedit para AutoEjecutarse
    limpiando_llave_de_inicio_del_virus
    mostrar_systray_notificacion LoadResString(33), LoadResString(26)
End If
'=========================================================
' Win2x
'=========================================================
If FileExist(Environ("windir") & "\system32\Win2x.exe") Then
    Dim Win2x As String
    Win2x = Environ("windir") & "\system32\Win2x.exe"
    'Matamos al proceso
    Shell "taskkill /f /im Win2x.exe", vbHide
    'Le quitamos los atributos
    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Win2x & WinSeek.doble_comillas.Text, vbHide
    'Eliminamos la entrada que crea en Regedit para AutoEjecutarse
    limpiando_llave_de_inicio_del_virus
    mostrar_systray_notificacion LoadResString(33), LoadResString(87)
End If
'=========================================================
' Win2x SERVICIOS (save.exe)
'=========================================================
If FileExist(Environ("windir") & "\system32\save.exe") Then
    Dim Win2xServicio As String
    Win2xServicio = Environ("windir") & "\system32\save.exe"
    'Matamos al proceso
    Shell "taskkill /f /im save.exe", vbHide
    'Le quitamos los atributos
    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Win2xServicio & WinSeek.doble_comillas.Text, vbHide
    'Eliminamos el archivo
    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & Win2xServicio & WinSeek.doble_comillas, vbHide
    'Eliminamos la entrada que crea en Regedit para AutoEjecutarse
    limpiando_llave_de_servicio_de_win2x
    mostrar_systray_notificacion LoadResString(33), LoadResString(90)
End If
End Sub

Public Function Obtener_Nombre_Archivo(p As String)
Dim Buffer As String
'Buffer de caracteres
Buffer = String(255, 0)

'Llamada a GetFileTitle, pasandole el path, el buffer y el tamaño
GetFileTitle p, Buffer, Len(Buffer)

'Retornamos el nombre eliminando los espacios nulos
Obtener_Nombre_Archivo = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)

End Function
