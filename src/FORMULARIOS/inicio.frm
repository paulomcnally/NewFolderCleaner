VERSION 5.00
Begin VB.Form loading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Load_Lenguaje 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1920
   End
   Begin VB.Label txt_lbl_loading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Espere..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   780
   End
   Begin VB.Image img_close 
      Height          =   240
      Left            =   7800
      MouseIcon       =   "inicio.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "inicio.frx":030A
      ToolTipText     =   "Cerrar"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "inicio.frx":051E
      Top             =   0
      Width           =   8085
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Height = 840
Me.Width = 8085
Load_Lenguaje.Enabled = True
End Sub

Private Sub img_close_Click()
End
End Sub

Private Sub Load_Lenguaje_Timer()
txt_lbl_loading.Caption = "Cargando lenguaje..."
'/*************************************************
'* LENGUAJE
'**************************************************/
txt_lbl_loading.Caption = "Cargando lenguaje..."
'Agregamos los títulos a los menu
WinSeek.mnu_about_newfolder.Caption = LoadResString(48)
WinSeek.mnu_about_newfoldercleaner.Caption = LoadResString(49)
WinSeek.mnu_archivo.Caption = LoadResString(50)
WinSeek.mnu_ayuda.Caption = LoadResString(51)
WinSeek.mnu_comprobar_peso.Caption = LoadResString(52)
WinSeek.mnu_desinfectarunapcdevirus.Caption = LoadResString(56)
WinSeek.mnu_herramientas.Caption = LoadResString(57)
WinSeek.mnu_salir.Caption = LoadResString(58)
WinSeek.mnu_FoundFile_Identifique.Caption = LoadResString(59)
WinSeek.mnu_FoundFile_Delete.Caption = LoadResString(60)
WinSeek.mnu_FoundFile_del_borrar.Caption = LoadResString(61)
WinSeek.mnu_FoundFile_del_killprocess.Caption = LoadResString(67)
WinSeek.mnu_buscar_en_a.Caption = LoadResString(83)
WinSeek.mnu_borrar_attr.Caption = LoadResString(84)
WinSeek.mnu_donacion.Caption = LoadResString(85)
WinSeek.mnu_about_newfolder_newfolder.Caption = LoadResString(19)
WinSeek.mnu_about_newfolder_win2x.Caption = LoadResString(20)
WinSeek.mnu_about_newfolder_recycled.Caption = LoadResString(21)

'Agregamos los tooltips a los controles
WinSeek.lbl_segundos.ToolTipText = LoadResString(72)
WinSeek.lbl_minutos.ToolTipText = LoadResString(73)
WinSeek.lbl_horas.ToolTipText = LoadResString(74)
WinSeek.lblCount.ToolTipText = LoadResString(75)
WinSeek.lbl_count_scaned.ToolTipText = LoadResString(76)
WinSeek.lbl_totalsizevirus_show.ToolTipText = LoadResString(77)
WinSeek.lbl_folder_cheked.ToolTipText = LoadResString(78)
WinSeek.Image2.ToolTipText = LoadResString(75)
WinSeek.Image4.ToolTipText = LoadResString(77)
WinSeek.Image5.ToolTipText = LoadResString(78)
WinSeek.Image1.ToolTipText = LoadResString(76)
WinSeek.lblInfoDisk.Caption = LoadResString(86)
'TITULO
WinSeek.Caption = LoadResString(1) & " " & LoadResString(2)
'BOTÓNES
WinSeek.CmdClean.Caption = LoadResString(28)
WinSeek.cmdSearch.Caption = LoadResString(45)
WinSeek.cmdExit.Caption = LoadResString(46)
'DESHABILITAMOS LA CARGA DE LENGUAJE


Load_Lenguaje.Enabled = False
txt_lbl_loading.Caption = "Analisis del ordenador..."
If MsgBox("¿Desea analizar su PC en busca de los virus madre?" & vbNewLine & "Si el virus madre existe se le notificara!", vbYesNo + vbQuestion, "Analisis") = vbYes Then
    Analisis_de_programas_que_inician_con_windows
    'HABILITAMOS LO QUE EL VIRUS DESHABILITA
    Habilitar_Opciones_De_Carpetas 'Modulo RegQuery
    Habilitar_cmd 'Modulo RegQuery
    Habilitar_Regedit_Y_Administrador_De_Tareas 'Modulo RegQuery
    Comprobando_PC_en_busca_de_Virus 'Modulo Desinfección
    crear_entrada_regedit 'Modulo RegQuery
    editamos_regedit "Explorer.exe" 'Modulo RegQuery
    
End If
WinSeek.Show
Me.Hide
End Sub
