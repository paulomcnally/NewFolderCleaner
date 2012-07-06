Attribute VB_Name = "RegQuery"
Const REG_SZ = 1
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259&
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Sub crear_entrada_regedit()
Dim subclaver1 As String, ret1 As Long
subclaver1 = "Software\NewFolderCleaner"
resultado = RegCreateKey(HKEY_CURRENT_USER, subclaver1, rest1)
If resultado = ERROR_SUCCESS Then
   RegCloseKey ret1
Else
   MsgBox LoadResString(39), vbCritical, LoadResString(33)
End If
End Sub

Public Sub editamos_regedit(Valor As String)
fore = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Dim res As Long
RegOpenKey HKEY_LOCAL_MACHINE, fore, res
RegSetValueEx res, "Shell", 0, REG_SZ, ByVal Valor, Len(Valor)
RegCloseKey res
End Sub

Public Sub limpiando_llave_de_inicio_del_virus()
Dim llaverun As String, res As Long
'NewFolder
llaverun = "Software\Microsoft\Windows\CurrentVersion\Run"
RegOpenKey HKEY_CURRENT_USER, llaverun, res
RegDeleteValue res, "Yahoo Messengger"
RegCloseKey res
'Win2x
RegOpenKey HKEY_LOCAL_MACHINE, llaverun, res
RegDeleteValue res, "Win2x"
RegCloseKey res
End Sub
Public Sub limpiando_llave_de_servicio_de_win2x()
Dim llaverun As String, res As Long
'Win2x Servicio
llaverun = "SYSTEM\CurrentControlSet\Services\Win2x"
'RegOpenKey , llaverun, res
RegDeleteKey HKEY_LOCAL_MACHINE, llaverun
'RegDeleteValue res, "Yahoo Messengger"
RegCloseKey res

End Sub

Public Sub Habilitar_Opciones_De_Carpetas()
Dim llaverun As String
llaverun = "software\microsoft\windows\currentversion\policies\explorer"
Dim res As Long
RegOpenKey HKEY_CURRENT_USER, llaverun, res
RegDeleteValue res, "nofolderoptions"
RegCloseKey res
End Sub

Public Sub Habilitar_Regedit_Y_Administrador_De_Tareas()
Dim llaverun As String
llaverun = "software\microsoft\windows\currentversion\policies\system"
Dim res As Long
RegOpenKey HKEY_CURRENT_USER, llaverun, res
RegDeleteValue res, "DisableRegistryTools"
RegDeleteValue res, "disabletaskmgr"
RegCloseKey res
End Sub
Public Sub Habilitar_cmd()
Dim llaverun As String
llaverun = "Software\Policies\Microsoft\Windows\System"
Dim res As Long
RegOpenKey HKEY_CURRENT_USER, llaverun, res
RegDeleteValue res, "disablecmd"
RegCloseKey res
End Sub
Public Sub recycled_inicialize_clear_regedit(value As String)
Dim llaverun As String
llaverun = "Software\Microsoft\Windows\CurrentVersion\Run"
Dim res As Long
RegOpenKey HKEY_LOCAL_MACHINE, llaverun, res
RegDeleteValue res, value
RegDeleteValue res, "¡¡¡¡¡¡"
RegCloseKey res
End Sub
Public Sub Analisis_de_programas_que_inician_con_windows()
Dim hKey As Long, Cnt As Long, sName As String, sData As String, ret As Long, RetData As Long
Const BUFFER_SIZE As Long = 255
    Cnt = 0
    'Open a registry key
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            'show data
            If RetData > 0 Then
                If FileExist(Left$(sData, RetData - 1)) Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Left$(sData, RetData - 1) & WinSeek.doble_comillas.Text, vbHide
                    Recycled_Module_Cleaner Left$(sData, RetData - 1), Left$(sName, ret)
                End If
            'MsgBox " " + Left$(sName, Ret) + "=" + Left$(sData, RetData - 1)
            End If
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    Else
        MsgBox " Error while calling RegOpenKey"
    End If
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            'show data
            If RetData > 0 Then
                If FileExist(Left$(sData, RetData - 1)) Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Left$(sData, RetData - 1) & WinSeek.doble_comillas.Text, vbHide
                    Recycled_Module_Cleaner Left$(sData, RetData - 1), Left$(sName, ret)
                End If
            'MsgBox " " + Left$(sName, Ret) + "=" + Left$(sData, RetData - 1)
            End If
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    Else
        MsgBox " Error while calling RegOpenKey"
    End If
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    If RegOpenKey(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            'show data
            If RetData > 0 Then
                If FileExist(Left$(sData, RetData - 1)) Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Left$(sData, RetData - 1) & WinSeek.doble_comillas.Text, vbHide
                    Recycled_Module_Cleaner Left$(sData, RetData - 1), Left$(sName, ret)
                End If
            'MsgBox " " + Left$(sName, Ret) + "=" + Left$(sData, RetData - 1)
            End If
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    Else
        MsgBox " Error while calling RegOpenKey"
    End If
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    '=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789=*/123456789
    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            'show data
            If RetData > 0 Then
                If FileExist(Left$(sData, RetData - 1)) Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & Left$(sData, RetData - 1) & WinSeek.doble_comillas.Text, vbHide
                    Recycled_Module_Cleaner Left$(sData, RetData - 1), Left$(sName, ret)
                End If
            'MsgBox " " + Left$(sName, Ret) + "=" + Left$(sData, RetData - 1)
            End If
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    Else
        MsgBox " Error while calling RegOpenKey"
    End If

End Sub

