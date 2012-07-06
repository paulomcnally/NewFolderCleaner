Attribute VB_Name = "ObtenerSizeyFecha"
Option Explicit
Public Const OFN_EXPLORER = &H80000

Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Sub getsizeofonefile()
Dim OFN As OPENFILENAME, Ret As Long
'Le establecemos valores a la estructura del cuadro de dialogo
With OFN
    .lStructSize = Len(OFN)
    .hInstance = App.hInstance
    .hwndOwner = WinSeek.hwnd
    .lpstrTitle = "Abrir archivo"
    .lpstrFilter = "EXE file (*.exe)" + Chr$(0) + "*.exe"
    .lpstrFile = String(255, 0)
    .nMaxFile = 255
    .flags = OFN_EXPLORER
End With

Ret = GetOpenFileName(OFN)
    
If Ret <> 0 Then
    CloseHandle Ret
    GeneralGetFileSizeAndDate (Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile, Chr$(0)) - 1))
End If
End Sub

