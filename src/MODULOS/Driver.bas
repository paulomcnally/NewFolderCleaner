Attribute VB_Name = "Driver"
'Función Api getLogicalDrives para recuperar los nombres de las unidades
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Función Api getLogicalDrives para recuperar las unidades
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
'SABER QUE TIPO DE UNIDAD ES
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

' Mensaje para SendMessage para establecer el control como Solo Lectura
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETSEL = &HB1

' Obtiene el handle o Hwnd del Edit del combobox
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'ESPACIO EN DISCO
Const PI = 3.14159

Const Kb As Double = 1024
Const Mb As Double = 1024 * Kb
Const Gb As Double = 1024 * Mb
Const Tb As Double = 1024 * Gb
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Type T_Info

    Capacidad As String
    Libre As String
    Usado As String
    
    CapacidadBytes As Double
    LibreBytes     As Double
    UsadoBytes     As Double
End Type
Public Sub add_driver()
Locked_Combo WinSeek.Combo_Drive_Load, True
Dim LDs As Long, Cnt As Long, sDrives As String

LDs = GetLogicalDrives
For Cnt = 0 To 25
    If (LDs And 2 ^ Cnt) <> 0 Then
        sDrives = Chr$(65 + Cnt) & ":\"
        Get_Drive_Volumen sDrives
        Select Case GetDriveType(sDrives)
         Case 2
             ' Unidad de tipo removible, por ejemplo la unidad A:
            If sDrives = "A:\" Then
            Else
                WinSeek.Combo_Drive_Load.AddItem sDrives & " [" & Get_Drive_Volumen(sDrives) & "]"
            End If
        Case 3
             ' Por ejemplo un disco duro
             WinSeek.Combo_Drive_Load.AddItem sDrives & " [" & Get_Drive_Volumen(sDrives) & "]"
              
         Case Is = 4
             'Una unidad de red
             'WinSeek.Combo_Drive_Load.AddItem sDrives
         Case Is = 5
             'Unidad de Cd - Dvd
            'WinSeek.Combo_Drive_Load.AddItem Left(sDrives, 3)
         Case Is = 6
             'Para una unidad de dsico Virtual
             WinSeek.Combo_Drive_Load.AddItem sDrives & " [" & Get_Drive_Volumen(sDrives) & "]"
         Case Else
             'Cuando es desconocida
             'WinSeek.Combo_Drive_Load.AddItem Left(sDrives, 3)
         End Select
     
    End If
       
Next Cnt
End Sub

Public Sub Locked_Combo(El_Combo As ComboBox, Bloquear As Boolean)
    Dim ret As Long, ret2 As Long
      
    Dim Hwnd_Combo As Long
      
    ' Obtiene el Hwnd del Edit del combobox
    Hwnd_Combo = FindWindowEx(El_Combo.hwnd, 0&, vbNullString, vbNullString)
      
    ' Le pasa el Hwnd del Edit del combobox y el mensaje EM_SETREADONLY _
      para establecerle el Locked, y el valor True o False para habilitar _
      o deshabilitarlo
      
    ret = SendMessage(Hwnd_Combo, EM_SETREADONLY, Bloquear, ByVal 0&)
  
 End Sub
Private Function Entero_a_Double(l As Long, h As Long) As Double

Dim ret As Double

    ret = h
    If h < 0 Then ret = ret + 2 ^ 32
    ret = ret * 2 ^ 32

    ret = ret + l
    If l < 0 Then ret = ret + 2 ^ 32

    Entero_a_Double = ret
End Function

Private Function Size(ByVal n_bytes As Double) As String

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

Function getInfoDrive(Drive As String) As T_Info

On Error GoTo errSub
    
    Dim Avalables As LARGE_INTEGER, Total As LARGE_INTEGER
    Dim Libres As LARGE_INTEGER, dTotal As Double, dLibre As Double
    Dim ret As Long
    
    ret = GetDiskFreeSpaceEx(Drive, Avalables, Total, Libres)
    
    dTotal = Entero_a_Double(Total.lowpart, Total.highpart)
    dLibre = Entero_a_Double(Libres.lowpart, Libres.highpart)

    ' retorna a la función los valores convertidos a String
    With getInfoDrive
        
        ' bytes
        .CapacidadBytes = dTotal
        .LibreBytes = dLibre
        
        ' string
        .Capacidad = Size(dTotal)
        .Libre = Size(dLibre)
        .Usado = Size(dTotal - dLibre)
    End With
    
Exit Function

'Error
errSub:
MsgBox Err.Description, vbCritical

End Function

Sub Dibujar_Circulo( _
    Valor_Maximo As Double, _
    Valor As Double, _
    Radio As Integer, _
    BackColor As Long, _
    ForeColor As Long, _
    ValueColor As Long, _
    BorderColor As Long, _
    Objeto As Object)
    
    If Valor_Maximo <= 0 Then Exit Sub
    
    With Objeto
        .BackColor = BackColor
        .ScaleMode = vbPixels
        Objeto.Cls

        Dim I As Long, per, xs, ys, cx, cy

        per = Valor / Valor_Maximo * 100
        per = per / 100
        per = 360 * per
    
        cx = .ScaleWidth \ 2
        cy = .ScaleHeight \ 2

        .DrawWidth = 2
    End With

    For I = 0 To 360
        xs = Cos(I / 180 * PI) * Radio
        ys = Sin(I / 180 * PI) * Radio
        Objeto.Line (cx, cy)-(cx + xs, cy + ys), ForeColor
        DoEvents
    Next I

    For I = 0 To per
        xs = Cos(I / 180 * PI) * Radio
        ys = Sin(I / 180 * PI) * Radio
        Objeto.Line (cx, cy)-(cx + xs, cy + ys), ValueColor
        DoEvents
    Next I
    
    With Objeto
        .DrawWidth = 2
        Objeto.Circle (.ScaleWidth / 2, .ScaleHeight / 2), Radio + 4, BorderColor
    End With
    
End Sub

Public Function Get_Drive_Volumen(ByVal s_Drive As String) As String
         
           Dim o_Fso As Scripting.FileSystemObject
           Dim o_Drive As Drive
             
           ' Creamos un nuevo objeto de tipo Scripting FileSystemObject
           Set o_Fso = New Scripting.FileSystemObject
             
           ' Si el Drive no es un vbnullstring
           If s_Drive <> "" Then
               ' Recuperamos el Drive para poder acceder _
                en las siguientes lineas
               Set o_Drive = o_Fso.GetDrive(s_Drive)
           End If
             
           With o_Drive
                If .IsReady Then
                  Get_Drive_Volumen = .VolumeName
                Else
                
                End If

           End With
             
           ' Eliminamos los objetos instanciados
           Set o_Drive = Nothing
           Set o_Fso = Nothing
             
End Function


