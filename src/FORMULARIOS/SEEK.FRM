VERSION 5.00
Begin VB.Form WinSeek 
   Caption         =   "WinSeek"
   ClientHeight    =   6105
   ClientLeft      =   1995
   ClientTop       =   2265
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   Icon            =   "SEEK.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_Combo 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8520
      Top             =   2880
   End
   Begin VB.CommandButton cmd_combo 
      Caption         =   "cmd_combo_Focus"
      Height          =   495
      Left            =   9000
      TabIndex        =   27
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox space_picture 
      AutoSize        =   -1  'True
      Height          =   1695
      Left            =   9120
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   25
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox doble_comillas 
      Height          =   285
      Left            =   9360
      TabIndex        =   20
      Text            =   """"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Picture_icon 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   9480
      Picture         =   "SEEK.frx":617A
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   17
      Top             =   5760
      Width           =   780
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4920
      TabIndex        =   16
      Top             =   3560
      Width           =   1440
   End
   Begin VB.Timer transcurrido 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9480
      Top             =   4320
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   8055
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ListBox lstFoundFiles 
         Height          =   2010
         ItemData        =   "SEEK.frx":15E84
         Left            =   120
         List            =   "SEEK.frx":15E86
         TabIndex        =   8
         Top             =   480
         Width           =   7815
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   3120
         Picture         =   "SEEK.frx":15E88
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lbl_folder_cheked 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   3480
         TabIndex        =   23
         Top             =   120
         Width           =   120
      End
      Begin VB.Label lbl_totalsizevirus_show 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   120
         Width           =   120
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1440
         Picture         =   "SEEK.frx":162A7
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   6360
         Picture         =   "SEEK.frx":166C8
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "SEEK.frx":1690D
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4680
         Picture         =   "SEEK.frx":16B6E
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lbl_count_scaned 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   120
         Width           =   120
      End
      Begin VB.Label lbl_segundos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   7680
         TabIndex        =   14
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Left            =   7560
         TabIndex        =   13
         Top             =   120
         Width           =   75
      End
      Begin VB.Label lbl_minutos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   7200
         TabIndex        =   12
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Left            =   7080
         TabIndex        =   11
         Top             =   120
         Width           =   75
      End
      Begin VB.Label lbl_horas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   120
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   120
      End
   End
   Begin VB.CommandButton CmdClean 
      Appearance      =   0  'Flat
      Caption         =   "&Eliminar Virus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3560
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   8055
      TabIndex        =   1
      Top             =   840
      Width           =   8055
      Begin VB.ComboBox Combo_Drive_Load 
         Height          =   315
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   90
         Width           =   2295
      End
      Begin VB.DirListBox dirList 
         Height          =   2115
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   7815
      End
      Begin VB.FileListBox filList 
         Height          =   2040
         Left            =   7920
         TabIndex        =   4
         Top             =   2760
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtSearchSpec 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Text            =   "*.exe"
         Top             =   2040
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label lblInfoDisk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "String(86)"
         Height          =   195
         Left            =   2520
         TabIndex        =   26
         Top             =   150
         Width           =   840
      End
      Begin VB.Label lblCriteria 
         Caption         =   "&Criterio de búsqueda:"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6480
      TabIndex        =   0
      Top             =   3560
      Width           =   1440
   End
   Begin VB.Label lbl_totalsizevirus 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   9360
      TabIndex        =   21
      Top             =   960
      Width           =   120
   End
   Begin VB.Image DONATION_IMG 
      Height          =   450
      Left            =   120
      MouseIcon       =   "SEEK.frx":16DD1
      MousePointer    =   99  'Custom
      Picture         =   "SEEK.frx":170DB
      ToolTipText     =   "Hacer una donación"
      Top             =   3490
      Width           =   795
   End
   Begin VB.Image BANNER_2 
      Height          =   840
      Left            =   -10
      Top             =   0
      Width           =   8085
   End
   Begin VB.Label lbl_file_revisados 
      AutoSize        =   -1  'True
      Caption         =   "Revisados:"
      Height          =   195
      Left            =   9480
      TabIndex        =   18
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label lbl_v_version 
      AutoSize        =   -1  'True
      Caption         =   "Versión del virus:"
      Height          =   195
      Left            =   9360
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Menu mnu_archivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnu_salir 
         Caption         =   "String(58)"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnu_herramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnu_comprobar_peso 
         Caption         =   "String(52)"
      End
      Begin VB.Menu mnu_borrar_attr 
         Caption         =   "String(84)"
      End
      Begin VB.Menu mnu_buscar_en_a 
         Caption         =   "String(83)"
      End
   End
   Begin VB.Menu mnu_ayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnu_donacion 
         Caption         =   "String(85)"
      End
      Begin VB.Menu line_ayuda_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_desinfectarunapcdevirus 
         Caption         =   "String(56)"
      End
      Begin VB.Menu line_ayuda_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about_newfolder 
         Caption         =   "String(48)"
         Begin VB.Menu mnu_about_newfolder_newfolder 
            Caption         =   "String(19)"
         End
         Begin VB.Menu mnu_about_newfolder_win2x 
            Caption         =   "String(20)"
         End
         Begin VB.Menu mnu_about_newfolder_recycled 
            Caption         =   "String(21)"
         End
      End
      Begin VB.Menu mnu_about_newfoldercleaner 
         Caption         =   "String(49)"
      End
   End
   Begin VB.Menu mnu_lstFoundFiles 
      Caption         =   "FoundFile"
      Visible         =   0   'False
      Begin VB.Menu mnu_FoundFile_Identifique 
         Caption         =   "String(59) - Identificar"
      End
      Begin VB.Menu mnu_FoundFile_Delete 
         Caption         =   "String(60) - Quitar"
      End
      Begin VB.Menu mnu_FoundFile_del_borrar 
         Caption         =   "String(61) - Borrar"
      End
      Begin VB.Menu mnu_FoundFile_del_killprocess 
         Caption         =   "Borrar este archivo [FORZADA]"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SearchFlag As Integer   ' Se usa como indicador para cancelar y para
                            ' otras operaciones.

Private Sub BANNER_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub BANNER_2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240

End Sub



Private Sub CmdClean_Click()
MakeCleanVirus
End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = LoadResString(46) Then
    
    Unload Me
    Else                    ' Si el usuario elige "Cancelar", acabar la
                            ' búsqueda.
        SearchFlag = False
    End If
End Sub

Private Sub cmdSearch_Click()
' Se inicializa para buscar, realizando una búsqueda recursiva.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Verifica lo último que hizo el usuario.
    If cmdSearch.Caption = LoadResString(44) Then  ' Si sólo fue "Restablecer", inicializa y sale.
        ResetSearch
        'txtSearchSpec.SetFocus
        Exit Sub
    End If

    ' Actualiza dirList.Path si es diferente del directorio seleccionado
    ' actualmente, si no es así realiza la búsqueda.
    If dirList.path <> dirList.List(dirList.ListIndex) Then
        dirList.path = dirList.List(dirList.ListIndex)
        Exit Sub         ' Sale para que el usuario pueda verificar antes de buscar.
    End If

    ' Continua con la búsqueda.
    Picture2.Move 0, 840
    Picture1.Visible = False
    Picture2.Visible = True
    lbl_segundos.Caption = "0"
    lbl_minutos.Caption = "0"
    lbl_horas.Caption = "0"
    transcurrido.Enabled = True
    
    cmdExit.Caption = LoadResString(47)

    filList.Pattern = txtSearchSpec.Text
    FirstPath = dirList.path
    DirCount = dirList.ListCount

    ' Inicia la búsqueda recursiva por directorios.
    
    NumFiles = 0                       ' Restablece el indicador de archivos encontrados.
    result = DirDiver(FirstPath, DirCount, "")
    filList.path = dirList.path
    cmdSearch.Caption = LoadResString(44)
    
    If lblCount > 0 Then
        Const Kb As Double = 1024
        Const Mb As Double = 1024 * Kb
        Const Gb As Double = 1024 * Mb
        Const Tb As Double = 1024 * Gb
        transcurrido.Enabled = False
        
        mostrar_systray_notificacion LoadResString(35), _
        LoadResString(69) & " " & lbl_count_scaned.Caption & " " & LoadResString(70) & " " & lbl_folder_cheked.Caption & " " & LoadResString(79) & " " & lblCount & LoadResString(36) & vbNewLine & LoadResString(71) & " " & Size(lbl_totalsizevirus)
        CmdClean.Enabled = True
        lbl_totalsizevirus = 0
    Else
        lbl_file_revisados.Visible = False
        transcurrido.Enabled = False
        mostrar_systray_notificacion LoadResString(35), _
        LoadResString(69) & " " & lbl_count_scaned.Caption & " " & LoadResString(70) & " " & lbl_folder_cheked.Caption & " " & LoadResString(79) & " " & lblCount & LoadResString(36)
        ResetSearch
        lbl_count_scaned.Caption = 0
    
    End If
    cmdExit.Caption = LoadResString(46)
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Buscar recursivamente en los directorios desde NewPath hacia abajo...
'  Esta búsqueda recursiva busca NewPath.
'  BackUp es el origen de la búsqueda recursiva.
'  DirCount es el número de subdirectorios en este directorio.
lbl_file_revisados.Visible = True
lbl_folder_cheked = Val(lbl_folder_cheked) + 1 'FoldesScanned
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Establece el indicador para que el usuario pueda interrumpir.
    DirDiver = False            ' Cambia a True si hay un error.
    retval = DoEvents()         ' Comprueba los eventos (por ejemplo, si el usuario elige "Cancelar").
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    
    DirsToPeek = dirList.ListCount                  ' ¿Cuantos directorios hay debajo de éste?
    Do While DirsToPeek > 0 And SearchFlag = True
    
        OldPath = dirList.path                      ' Guardar la ruta antigua para la próxima búsqueda recursiva.
        dirList.path = NewPath
        If dirList.ListCount > 0 Then
        
            ' Ir al nodo inferior.
            dirList.path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.path), DirCount%, OldPath)
        End If
        ' Sube un nivel de directorios.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Llamar a la función para enumerar archivos.
    If filList.ListCount Then
        If Len(dirList.path) <= 3 Then             ' Buscar caracteres de dos bytes
            ThePath = dirList.path                  ' Si es en el nivel raíz, dejarlo como está...
        Else
            ThePath = dirList.path + "\"            ' si no es así poner "\" antes del archivo.
        End If
        For ind = 0 To filList.ListCount - 1        ' Agregar los archivos adecuados en este directorio al cuadro de lista.
            entry = ThePath + filList.List(ind)
            'Buscamos y eliminamos amenaza
            GetDirSize entry, False
        
        Next ind
        
    End If
    
    If BackUp <> "" Then        ' Si hay un directorio superior, moverlo.
        dirList.path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' Si ocurre un error de memoria insuficiente, asumir que el cuadro de lista se ha llenado.
        DirDiver = True         ' Crear un mensaje y establecer el valor de retorno AbandonSearch.
        MsgBox LoadResString(34), vbInformation, LoadResString(33)
        Exit Function           ' Observar que el procedimiento de salida restablece Err a 0.
    Else                        ' Si no es así muestra el mensaje de error y sale.
        MsgBox Error
        End
    End If
    
End Function


Private Sub Combo_Drive_Load_Click()
dirList.path = Left(Combo_Drive_Load.Text, 3)

'AGREGARMOS LA INFO DE LA UNIDAD
Dim tDrive As T_Info
    ' obtener info
    tDrive = getInfoDrive(Left(Combo_Drive_Load.Text, 3))
    
    ' dibujar
    With tDrive
        
        Call Dibujar_Circulo( _
            .CapacidadBytes, _
            .LibreBytes, 50, _
            vbBlack, vbGreen, _
            RGB(52, 114, 197), vbBlack, space_picture)
        
        ' captions
        lblInfoDisk.Caption = "Capacidad: " & .Capacidad & " | " & "Libre: " & .Libre & " | " & "Usada: " & .Usado
    End With
Timer_Combo.Enabled = True
End Sub

Private Sub Combo_Drive_Load_gotFocus()
Dim ret2 As Long
ret2 = SendMessageLong(Combo_Drive_Load.hwnd, &H14F, True, 0)
End Sub

Private Sub Combo_Drive_Load_Validate(Cancel As Boolean)
doble_comillas.SetFocus
End Sub

Private Sub DirList_Change()
    ' Actualizar el cuadro de lista de archivos para sincronizarlo con el
    ' cuadro de lista del directorio.
    filList.path = dirList.path
End Sub

Private Sub DirList_LostFocus()
    dirList.path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DONATION_IMG_Click()
Shell Environ("programfiles") & "\Internet Explorer\iexplore.exe " & "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=842XBFR63HNFQ&lc=GB&item_name=Polin&item_number=Polin&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG_global%2egif%3aNonHosted", vbMaximizedFocus
End Sub

Private Sub Form_Load()

'RERRENAMOS EL COMBO DE LOS DRIVES
add_driver 'Modulo Driver
If Combo_Drive_Load.ListCount <> 0 Then
    Combo_Drive_Load.Text = Combo_Drive_Load.List(0)
    dirList.path = Left(Combo_Drive_Load.Text, 3)
End If


'PROPIEDADES DEL FORMULARIO
Picture_icon.Picture = Me.Icon
Propiedades_del_formulario Me, True, True, False, 4875, 8310 'Modulo apariencia
'IMAGEN DE BANNER
BANNER_2.Picture = LoadResPicture("POLIN2", vbResBitmap)

'MOSTRAR SIS TRAY
 With sysTray
         .cbSize = Len(sysTray)
         .hwnd = Me.Picture_icon.hwnd
         .uID = vbNull
         .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Picture_icon.Picture
         .szTip = LoadResString(1) & " " & LoadResString(2) & vbNullChar
         .dwState = 0
         .dwStateMask = 0
End With
'Notifica en el systray
Shell_NotifyIcon NIM_ADD, sysTray
'TRANSPARENCIA
aplicar_transparencia Me.hwnd, 240
'APARIENCIA DE LOS CARGADORES
    Picture2.Move 0, 840
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
    lblCriteria.BackColor = WinSeek.BackColor

    Picture1.Move 0, 840
    Picture1.Width = WinSeek.ScaleWidth
    Picture1.BackColor = WinSeek.BackColor


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Form_Resize()
'If (Me.WindowState = vbMinimized) Then
'mostrar_systray_notificacion "Hola", "Estamos aquí!!!"
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, sysTray
'NewFolder 1
If FileExist(Environ("windir") & "\RVHOST.exe") Then
    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & Environ("windir") & "\RVHOST.exe" & WinSeek.doble_comillas, vbHide
End If
'NewFolder 2
If FileExist(Environ("windir") & "\system32\RVHOST.exe") Then
    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & Environ("windir") & "\system32\RVHOST.exe" & WinSeek.doble_comillas, vbHide
End If
'Win2x
If FileExist(Environ("windir") & "\system32\Win2x.exe") Then
    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & Environ("windir") & "\system32\Win2x.exe" & WinSeek.doble_comillas, vbHide
End If
'Recycled
'Comprobamos si existe el archivo en cada unidad para eliminarle
Dim LDs As Long, Cnt As Long, sDrives As String

LDs = GetLogicalDrives
For Cnt = 0 To 25
    If (LDs And 2 ^ Cnt) <> 0 Then
        sDrives = Chr$(65 + Cnt) & ":\"
        'Get_Drive_Volumen sDrives
        Select Case GetDriveType(sDrives)
         Case 2
             ' Unidad de tipo removible, por ejemplo la unidad A:
            If sDrives = "A:\" Then
            Else
                If FileExist(sDrives & "Recycled.exe") Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "Recycled.exe" & WinSeek.doble_comillas.Text, vbHide
                    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "Recycled.exe" & WinSeek.doble_comillas, vbHide
                    'DeleteFile sDrives & "Recycled.exe"
                End If
                If FileExist(sDrives & "autorun.inf") Then
                    Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "autorun.inf.exe" & WinSeek.doble_comillas.Text, vbHide
                    Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "autorun.inf" & WinSeek.doble_comillas, vbHide
                End If
            End If
        Case 3
             ' Por ejemplo un disco duro
            If FileExist(sDrives & "Recycled.exe") Then
                Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "Recycled.exe" & WinSeek.doble_comillas.Text, vbHide
                Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "Recycled.exe" & WinSeek.doble_comillas, vbHide

            End If
            If FileExist(sDrives & "autorun.inf") Then
                Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "autorun.inf" & WinSeek.doble_comillas.Text, vbHide
                Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "autorun.inf" & WinSeek.doble_comillas, vbHide
            End If
         Case Is = 4
             'Una unidad de red
             'WinSeek.Combo_Drive_Load.AddItem sDrives
         Case Is = 5
             'Unidad de Cd - Dvd
            'WinSeek.Combo_Drive_Load.AddItem Left(sDrives, 3)
         Case Is = 6
             'Para una unidad de dsico Virtual
            If FileExist(sDrives & "Recycled.exe") Then
                Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "Recycled.exe" & WinSeek.doble_comillas.Text, vbHide
                Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "Recycled.exe" & WinSeek.doble_comillas, vbHide
            End If
            If FileExist(sDrives & "autorun.inf") Then
                Shell "attrib -r -a -s -h " & WinSeek.doble_comillas.Text & sDrives & "autorun.inf" & WinSeek.doble_comillas.Text, vbHide
                Shell "cmd /C del /a/f/q " & WinSeek.doble_comillas & sDrives & "autorun.inf" & WinSeek.doble_comillas, vbHide
            End If
         Case Else
             'Cuando es desconocida
             'WinSeek.Combo_Drive_Load.AddItem Left(sDrives, 3)
         End Select
     
    End If
       
Next Cnt
'Cierre de la aplicación
End
End Sub

Private Sub ResetSearch()
    ' Reinicializa antes de empezar una nueva búsqueda.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Indica que una búsqueda está en ejecución.
    Picture2.Visible = False
    cmdSearch.Caption = LoadResString(45)
    cmdExit.Caption = LoadResString(46)
    Picture1.Visible = True
    lbl_count_scaned.Caption = 0
    CmdClean.Enabled = False
    lbl_totalsizevirus_show = 0
    lbl_folder_cheked.Caption = 0
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_count_scaned_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_count_scaned_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_folder_cheked_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_folder_cheked_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_horas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_horas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_minutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_minutos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_segundos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_segundos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_version_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_version_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_totalsizevirus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_totalsizevirus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_totalsizevirus_show_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_totalsizevirus_show_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_v_version_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_v_version_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lblCount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lblCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lblfound_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub lblfound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub lblInfoDisk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lblInfoDisk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lstFoundFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstFoundFiles.SelCount > 0 Then
Call Show_Menu_PopUp(mnu_lstFoundFiles, Button)
End If
End Sub

Private Sub mnu_about_newfolder_newfolder_Click()
sub_mnu_about_newfolder_newfolder
End Sub

Private Sub mnu_about_newfolder_recycled_Click()
sub_mnu_about_newfolder_recycled
End Sub

Private Sub mnu_about_newfolder_win2x_Click()
sub_mnu_about_newfolder_win2x
End Sub

Private Sub mnu_about_newfoldercleaner_Click()
About.Show vbModal
End Sub


Private Sub mnu_borrar_attr_Click()
Dim ruta As String
ruta = InputBox("Escriba la ruta exácta del archivo o carpeta.", "Selección de archivo o carpeta")
If ruta = "" Then

Else
    Shell "attrib -r -a -s -h " & doble_comillas.Text & ruta & doble_comillas.Text, vbHide
    mostrar_systray_notificacion LoadResString(33), "Se quitaron los atributos a" & " " & ruta
End If
End Sub

Private Sub mnu_buscar_en_a_Click()
TIULO_MSGBOX = LoadResString(33)
SetTimer WinSeek.hwnd, 1, 1, AddressOf TimerProc
If MsgBox(LoadResString(81), vbOKCancel + vbInformation, TIULO_MSGBOX) = vbOK Then
KillTimer WinSeek.hwnd, 1
    On Error GoTo FALLAALCARGARUNIDADA:
    dirList.path = "A:\"
    cmdSearch.value = True
End If
FALLAALCARGARUNIDADA:
mostrar_systray_notificacion LoadResString(33), LoadResString(82)
Exit Sub
End Sub

Private Sub mnu_comprobar_peso_Click()
getsizeofonefile 'Modulo ObtenerSizeyFecha
End Sub

Private Sub mnu_consejosdeusodelaaplicacion_Click()
sub_mnu_consejosdeusodelaaplicacion
End Sub

Private Sub mnu_desabilitarmensajedeentrada_Click()
If mnu_desabilitarmensajedeentrada.Checked = True Then
mnu_desabilitarmensajedeentrada.Checked = False
editamos_regedit 0
Else
mnu_desabilitarmensajedeentrada.Checked = True
editamos_regedit 1
End If


End Sub

Private Sub mnu_desinfectarpcdelvirusnewfolder_Click()
Shell "tools/FlashDesinfector.exe", vbNormalFocus
End Sub

Private Sub mnu_desinfectarunapcdevirus_Click()
Shell Environ("programfiles") & "/Internet Explorer/iexplore.exe " & "http://polin.wareznica.com/index.php?accion=verlistavideotutoriales", vbMaximizedFocus
End Sub

Private Sub mnu_donacion_Click()
mostrar_systray_notificacion LoadResString(63), LoadResString(30)
Shell Environ("programfiles") & "\Internet Explorer\iexplore.exe " & LoadResString(62), vbMaximizedFocus
End Sub

Private Sub mnu_FoundFile_del_borrar_Click()
Dim seleccion As Long, ruta As String
seleccion = lstFoundFiles.ListIndex
ruta = lstFoundFiles.List(seleccion)
DeleteFile ruta
lstFoundFiles.RemoveItem (seleccion)
mostrar_systray_notificacion LoadResString(33), LoadResString(65)
If lstFoundFiles.ListCount = 0 Then
    CmdClean.Enabled = False
End If
End Sub

Private Sub mnu_FoundFile_del_killprocess_Click()
Dim seleccion As Long, patch As String, file As String, comando As String
seleccion = lstFoundFiles.ListIndex
patch = lstFoundFiles.List(seleccion)
file = Obtener_Nombre_Archivo(patch) & ".exe"
Shell "cmd /c taskkill /f /im " & doble_comillas.Text & file & doble_comillas.Text, vbHide
mostrar_systray_notificacion LoadResString(33), LoadResString(66)
End Sub

Private Sub mnu_FoundFile_Delete_Click()
Dim seleccion As Long
seleccion = lstFoundFiles.ListIndex
lstFoundFiles.RemoveItem (seleccion)
If lstFoundFiles.ListCount = 0 Then
    CmdClean.Enabled = False
End If
End Sub

Private Sub mnu_FoundFile_Identifique_Click()
Dim seleccion As Long, patch As String
seleccion = lstFoundFiles.ListIndex
patch = lstFoundFiles.List(seleccion)

GetDirSize patch, True
End Sub

Private Sub mnu_salir_Click()
Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Timer_Combo_Timer()
cmd_combo.SetFocus
'KeyAscii = 0
Timer_Combo.Enabled = False
End Sub

Private Sub transcurrido_Timer()
lbl_segundos.Caption = Val(lbl_segundos.Caption) + 1
If lbl_segundos.Caption = "60" Then
lbl_minutos.Caption = Val(lbl_minutos.Caption) + 1
lbl_segundos.Caption = "1"
End If
If lbl_minutos.Caption = "60" Then
lbl_horas.Caption = Val(lbl_horas.Caption) + 1
lbl_minutos.Caption = "1"
End If
End Sub

Private Sub txtSearchSpec_Change()
    ' Actualiza el cuadro de lista de archivos si el usuario cambia el modelo.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Resalta la entrada actual.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub




