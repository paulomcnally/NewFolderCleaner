Attribute VB_Name = "FileSize"
'ATRIBUTOS
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes    As Long
        ftCreationTime      As FILETIME
        ftLastAccessTime    As FILETIME
        ftLastWriteTime     As FILETIME
        nFileSizeHigh       As Long
        nFileSizeLow        As Long
        dwReserved0         As Long
        dwReserved1         As Long
        cFileName           As String * MAX_PATH
        cAlternate          As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'ATRIBUTOS / FIN

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'############################## FILETIME ######################################

'############################## GETFILESIZE ###################################

'Constante para pasar que indica que se abre el archivo en modo lectura
Public Const OF_READ = &H0&

' Api lOpen para abrir un archivo
Public Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long

' Api lclose para cerrar el archivo
Public Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

' Api GetFileSize para averiguar el tamaño
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public lpFSHigh As Long
Public Sub GetDirSize(path As String, show_msgbox As Boolean)
Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME, RSULT_FECHA As String
Dim Handle As Long
Dim RSULT_SIZE As Long
'----------------------------------------------------------------------
'CONSTANTES PARA LA VERSION 3.2.0.1 - NewFolder
Const SIZE_V_3201 = 268289
Const FECHA_V_3201 = 6182007
'CONSTANTES PARA LA VERSION 3.2.0.1_2 - NewFolder
Const SIZE_V_3201_2 = 309761
Const FECHA_V_3201_2 = 6182007
'CONSTANTES PARA LA VERSION 3.2.0.0 - NewFolder
Const SIZE_V_3220 = 312439
Const FECHA_V_3220 = 9262008
'CONSTANTES PARA Win2x 01
Const SIZE_Win2x_01 = 98304
Const FECHA_Win2x_01 = 5162007
'CONSTANTES PARA Recycled 01
Const SIZE_Recycled_01 = 1244127
'-----------------------------------------------------------------------
'Verifica la barra separadora de path
If Len(path) <= 3 Then
    path = path & "\"
Else
    path = path
End If
'ABRIMOS EL ARCHIVO EN LECTURA
Handle = lOpen(path, OF_READ)
'TAMAÑO DEL ARCHIVO
RSULT_SIZE = RSULT_SIZE + GetFileSize(Handle, lpFSHigh)
'FECHA DEL ARCHIVO
GetFileTime Handle, Ft1, Ft1, Ft2
FileTimeToLocalFileTime Ft2, Ft1
FileTimeToSystemTime Ft1, SysTime
RSULT_FECHA = Str$(SysTime.wMonth) + LTrim(Str$(SysTime.wDay)) + LTrim(Str$(SysTime.wYear))
'================================================================
'= NEW FOLDER ===================================================
'= VERSION 3.2.0.1 | SIZE 268289 | FECHA 6182007 ================
'= Sábado 15 de Noviembre del 2008 ==============================
'================================================================
If RSULT_SIZE = SIZE_V_3201 And RSULT_FECHA = FECHA_V_3201 Then
    If show_msgbox = True Then
        mostrar_systray_notificacion path, "Virus: NewFolder" & vbNewLine & _
        "Tamaño en byte: " & SIZE_V_3201 & vbNewLine & "Fecha de creación: " & FECHA_V_3201
        lclose Handle
        Exit Sub
    End If
    
    'AÑADIMOS EL PATCH AL LIST
    WinSeek.lstFoundFiles.AddItem path
    'AUMENTAMOS EL NUMERO DE VIRUS ENCONTRADOS
    WinSeek.lblCount.Caption = Str(Val(WinSeek.lblCount.Caption) + 1)
    'Sumamos el tamaño de cada archivo encontrado
    WinSeek.lbl_totalsizevirus = WinSeek.lbl_totalsizevirus + RSULT_SIZE
    WinSeek.lbl_totalsizevirus_show = Size(WinSeek.lbl_totalsizevirus)
'Cerramos el archivo que abrimos para la lectura de sus datos
lclose Handle
'Obtenemos el nombre + extención del archivo
Dim file As String
file = Obtener_Nombre_Archivo(path) & ".exe"
'Cerramos el proceso del archivo con taskkill
Shell "cmd /c taskkill /f /im " & WinSeek.doble_comillas.Text & file & WinSeek.doble_comillas.Text, vbHide
'Notificamos por medio de systray
mostrar_systray_notificacion LoadResString(33), "Cerrando Proceso: " & file
End If
'================================================================
'= NEW FOLDER ===================================================
'= VERSION 3.2.0.1 | SIZE 309761 | FECHA 6182007 ================
'= Miércoles 21 de Enero del 2009 ===============================
'================================================================
If RSULT_SIZE = SIZE_V_3201_2 And RSULT_FECHA = FECHA_V_3201_2 Then
    If show_msgbox = True Then
        mostrar_systray_notificacion path, "Virus: NewFolder" & vbNewLine & _
        "Tamaño en byte: " & SIZE_V_3201_2 & vbNewLine & "Fecha de creación: " & FECHA_V_3201_2 & _
        vbNewLine & "Versión notificada por Nestor Miranda. (NAMV) (Lunes 19 de Enero 2009)"
        lclose Handle
        Exit Sub
    End If
    
    'AÑADIMOS EL PATCH AL LIST
    WinSeek.lstFoundFiles.AddItem path
    'AUMENTAMOS EL NUMERO DE VIRUS ENCONTRADOS
    WinSeek.lblCount.Caption = Str(Val(WinSeek.lblCount.Caption) + 1)
    'Sumamos el tamaño de cada archivo encontrado
    WinSeek.lbl_totalsizevirus = WinSeek.lbl_totalsizevirus + RSULT_SIZE
    WinSeek.lbl_totalsizevirus_show = Size(WinSeek.lbl_totalsizevirus)
'Cerramos el archivo que abrimos para la lectura de sus datos
lclose Handle
'Obtenemos el nombre + extención del archivo
Dim file2 As String
file2 = Obtener_Nombre_Archivo(path) & ".exe"
'Cerramos el proceso del archivo con taskkill
Shell "cmd /c taskkill /f /im " & WinSeek.doble_comillas.Text & file2 & WinSeek.doble_comillas.Text, vbHide
'Notificamos por medio de systray
mostrar_systray_notificacion LoadResString(33), "Cerrando Proceso: " & file
End If
'================================================================
'= NEW FOLDER ===================================================
'= VERSION 3.2.2.0 | SIZE 312439 | FECHA 9262008 ================
'= Miércoles 21 de Enero del 2009 ===============================
'================================================================
If RSULT_SIZE = SIZE_V_3220 And RSULT_FECHA = FECHA_V_3220 Then
    If show_msgbox = True Then
        mostrar_systray_notificacion path, "Virus: NewFolder" & vbNewLine & _
        "Tamaño en byte: " & SIZE_V_3220 & vbNewLine & "Fecha de creación: " & FECHA_V_3220 & _
        vbNewLine & "Versión notificada por carlosmk66@gmail.com (Jueves 05 de Marzo 2009)"
        lclose Handle
        Exit Sub
    End If
    
    'AÑADIMOS EL PATCH AL LIST
    WinSeek.lstFoundFiles.AddItem path
    'AUMENTAMOS EL NUMERO DE VIRUS ENCONTRADOS
    WinSeek.lblCount.Caption = Str(Val(WinSeek.lblCount.Caption) + 1)
    'Sumamos el tamaño de cada archivo encontrado
    WinSeek.lbl_totalsizevirus = WinSeek.lbl_totalsizevirus + RSULT_SIZE
    WinSeek.lbl_totalsizevirus_show = Size(WinSeek.lbl_totalsizevirus)
'Cerramos el archivo que abrimos para la lectura de sus datos
lclose Handle
'Obtenemos el nombre + extención del archivo
Dim file5 As String
file5 = Obtener_Nombre_Archivo(path) & ".exe"
'Cerramos el proceso del archivo con taskkill
Shell "cmd /c taskkill /f /im " & WinSeek.doble_comillas.Text & file5 & WinSeek.doble_comillas.Text, vbHide
'Notificamos por medio de systray
mostrar_systray_notificacion LoadResString(33), "Cerrando Proceso: " & file
End If
'================================================================
'= Win2x ========================================================
'= VERSION 1 | SIZE 309761 | FECHA 6182007 ======================
'= Martes 27 de Enero del 2009 ===============================
'================================================================
If RSULT_SIZE = SIZE_Win2x_01 And RSULT_FECHA = FECHA_Win2x_01 Then
    If show_msgbox = True Then
        mostrar_systray_notificacion path, "Virus: Win2x" & vbNewLine & _
        "Tamaño en byte: " & SIZE_Win2x_01 & vbNewLine & "Fecha de creación: " & FECHA_Win2x_01 & vbNewLine & _
        "Versión notificada por Dayanara Patricia Gómez Ortíz (Martes 27 de Enero 2009)"
        lclose Handle
        Exit Sub
    End If
    
    'AÑADIMOS EL PATCH AL LIST
    WinSeek.lstFoundFiles.AddItem path
    'AUMENTAMOS EL NUMERO DE VIRUS ENCONTRADOS
    WinSeek.lblCount.Caption = Str(Val(WinSeek.lblCount.Caption) + 1)
    'Sumamos el tamaño de cada archivo encontrado
    WinSeek.lbl_totalsizevirus = WinSeek.lbl_totalsizevirus + RSULT_SIZE
    WinSeek.lbl_totalsizevirus_show = Size(WinSeek.lbl_totalsizevirus)
'Cerramos el archivo que abrimos para la lectura de sus datos
lclose Handle
'Obtenemos el nombre + extención del archivo
Dim file3 As String
file3 = Obtener_Nombre_Archivo(path) & ".exe"
'Cerramos el proceso del archivo con taskkill
Shell "cmd /c taskkill /f /im " & WinSeek.doble_comillas.Text & file2 & WinSeek.doble_comillas.Text, vbHide
'Notificamos por medio de systray
mostrar_systray_notificacion LoadResString(33), "Cerrando Proceso: " & file
End If
'================================================================
'= RECYCLED =====================================================
'= VERSION 1 | SIZE 1244127 =====================================
'= Sábado 24 de Enero del 2009 ==================================
'================================================================
If RSULT_SIZE = SIZE_Recycled_01 Then
    sizepath = Len(path)
    sizepathinext = sizepath - 4
    folderpatch = Left(path, sizepathinext)
    If FileExist(folderpatch) Then
        Dim atrib As Long
        atrib = GetAttr(folderpatch)
        'If atrib And vbHidden Then
            'Si se da al menu pop up que muestre la info
            If show_msgbox = True Then
            mostrar_systray_notificacion path, "Virus: Recycled" & vbNewLine & _
            "Tamaño en byte: " & 1244127
            lclose Handle
            Exit Sub
            End If
            'AÑADIMOS EL PATCH AL LIST
            WinSeek.lstFoundFiles.AddItem path
            'AUMENTAMOS EL NUMERO DE VIRUS ENCONTRADOS
            WinSeek.lblCount.Caption = Str(Val(WinSeek.lblCount.Caption) + 1)
            'Sumamos el tamaño de cada archivo encontrado
            WinSeek.lbl_totalsizevirus = WinSeek.lbl_totalsizevirus + RSULT_SIZE
            WinSeek.lbl_totalsizevirus_show = Size(WinSeek.lbl_totalsizevirus)
            'Cerramos el archivo que abrimos para la lectura de sus datos
            lclose Handle
            
            'Obtenemos el nombre + extención del archivo
            Dim file4 As String
            file4 = Obtener_Nombre_Archivo(path) & ".exe"
            
            'Cerramos el proceso del archivo con taskkill
            Shell "cmd /c taskkill /f /im " & WinSeek.doble_comillas.Text & file & WinSeek.doble_comillas.Text, vbHide
            'Notificamos por medio de systray
            mostrar_systray_notificacion LoadResString(33), "Cerrando Proceso: " & file
        'End If
    End If
End If
'==============================================================.
'
'==============================================================
mostrar_systray_notificacion LoadResString(29), WinSeek.lbl_file_revisados
WinSeek.lbl_file_revisados.Caption = path
WinSeek.lbl_file_revisados.Refresh
WinSeek.lbl_count_scaned.Caption = WinSeek.lbl_count_scaned.Caption + 1
End Sub

Public Sub GeneralGetFileSizeAndDate(dir As String)
'**********************Declaración de variables**************************************
Dim Handle As Long
Dim TamanioDelFile As Long
Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME, resultadov3201 As String
'**********************Leemos el archivo**************************************
Handle = lOpen(dir, OF_READ)
'**********************API GetFileSize**************************************
TamanioDelFile = TamanioDelFile + GetFileSize(Handle, lpFSHigh)
'**********************API GetFileTime**************************************
GetFileTime Handle, Ft1, Ft1, Ft2
FileTimeToLocalFileTime Ft2, Ft1
FileTimeToSystemTime Ft1, SysTime
FechaDelFile = Str$(SysTime.wMonth) + LTrim(Str$(SysTime.wDay)) + LTrim(Str$(SysTime.wYear))  'Imprimimos la fehca
'**********************Mostramos el mensaje**************************************
'TIULO_MSGBOX = LoadResString(42)
'SetTimer WinSeek.hWnd, 1, 1, AddressOf TimerProc
'MsgBox dir & vbNewLine & vbNewLine & TamanioDelFile & LoadResString(40) & vbNewLine & vbNewLine & FechaDelFile & LoadResString(41), vbInformation, TIULO_MSGBOX
'KillTimer WinSeek.hWnd, 1
'**********************Cerramos el archivo**************************************
mostrar_systray_notificacion "Resultado", dir & vbNewLine & vbNewLine & TamanioDelFile & LoadResString(40) & vbNewLine & vbNewLine & FechaDelFile & LoadResString(41)
lclose Handle
End Sub
Public Function Size(ByVal n_bytes As Double) As String
Const Kb As Double = 1024
Const Mb As Double = 1024 * Kb
Const Gb As Double = 1024 * Mb
Const Tb As Double = 1024 * Gb

If n_bytes < Kb Then
         Size = Format$(n_bytes) & " bytes"
ElseIf n_bytes < Mb Then
         Size = Format$(n_bytes / Kb, "0.00") & " KB"
ElseIf n_bytes < Gb Then
         Size = Format$(n_bytes / Mb, "0.00") & " MB"
Else
         Size = Format$(n_bytes / Gb, "0.00") & " GB"
End If
End Function

Public Function FileExist(ByVal sFile As String) As Boolean
    'comprobar si existe este fichero
    Dim WFD As WIN32_FIND_DATA
    Dim hFindFile As Long

    hFindFile = FindFirstFile(sFile, WFD)
    'Si no se ha encontrado
    If hFindFile = INVALID_HANDLE_VALUE Then
        FileExist = False
    Else
        FileExist = True
        'Cerrar el handle de FindFirst
        hFindFile = FindClose(hFindFile)
    End If

End Function

Public Sub Recycled_Module_Cleaner(dirrectorio As String, name As String)
'**********************Declaración de variables**************************************
Dim Handle As Long, SizeObtenido As Long, hKey As Long
Const size_recycled_1 = 1244127
'**********************Leemos el archivo**************************************
Handle = lOpen(dirrectorio, OF_READ)
'**********************API GetFileSize**************************************
SizeObtenido = SizeObtenido + GetFileSize(Handle, lpFSHigh)
'**********************API GetFileTime**************************************
If SizeObtenido = size_recycled_1 Then
'Matamos al proiceso del archivo
Shell "taskkill /f /im " & WinSeek.doble_comillas & name & ".EXE" & WinSeek.doble_comillas, vbHide
'Eliminamos Atributos al archivo
Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & dirrectorio & WinSeek.doble_comillas.Text, vbHide
'Eliminamos entrada en Regedit
recycled_inicialize_clear_regedit name
'Notificamos que hicimos limpieza del virus
mostrar_systray_notificacion LoadResString(33), LoadResString(88) & "(" & name & ".exe)" & LoadResString(89)
'Cerramos el archivo abierto
lclose Handle
'Eliminamos el archivo
Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & dirrectorio & WinSeek.doble_comillas, vbHide
Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & dirrectorio & WinSeek.doble_comillas, vbHide
End If
'MsgBox "PC infectada con virus Recycled! Click en Aceptar para Desinfectar la PC", vbCritical, "PC Infectada"
'
'Kill dirrectorio
End Sub
