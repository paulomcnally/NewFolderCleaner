Attribute VB_Name = "MakeClean"
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const LB_DELETESTRING = &H182

Public Sub MakeCleanVirus()

On Error GoTo ERRORDEELIMINACION:
For I = 0 To WinSeek.lblCount
    'Quitamos los atributos a las carpetas que el virus escondio'
    largo = Len(WinSeek.lstFoundFiles.List(0))
    sinextencion = largo - 4
    file = Left(WinSeek.lstFoundFiles.List(0), sinextencion)
    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & file & WinSeek.doble_comillas.Text, vbHide
    'Eliminamos el archivo y quitamos de la lista
    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & WinSeek.lstFoundFiles.List(0) & WinSeek.doble_comillas, vbHide
    'DeleteFile WinSeek.lstFoundFiles.List(0)
    mostrar_systray_notificacion LoadResString(80), WinSeek.lstFoundFiles.List(0)
    SendMessageLong WinSeek.lstFoundFiles.hwnd, LB_DELETESTRING, 0, 1
       
ERRORDEELIMINACION:
'MsgBox Error.Description, vbCritical, "ERROR"
Next I
WinSeek.lbl_count_scaned.Caption = 0
WinSeek.lblCount.Caption = 0
WinSeek.CmdClean.Enabled = False
mostrar_systray_notificacion LoadResString(33), LoadResString(38)
WinSeek.cmdSearch.value = True
WinSeek.lbl_totalsizevirus = 0
End Sub
