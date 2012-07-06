VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Titulo + Versión"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Licence_User 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label lbl_correo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "paulomcnally@yahoo.com"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Image Banner_IMAGE 
      Height          =   840
      Left            =   -10
      Top             =   0
      Width           =   4830
   End
   Begin VB.Label lbl_url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://polin.wareznica.com/index.php?accion=newfoldercleaner"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "About.frx":0000
      TabIndex        =   3
      ToolTipText     =   "Ir a la página Web de NewFolder Cleaner"
      Top             =   1920
      Width           =   4545
   End
   Begin VB.Label lbl_fecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label lbl_autor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paulo Antonio McNally Zambrana"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   2370
   End
   Begin VB.Label lbl_titleyversion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo + Versión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image icono 
      Height          =   615
      Left            =   240
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
     ByVal lpBuffer As String, _
     nSize As Long) As Long

Private Sub Banner_IMAGE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Banner_IMAGE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_url.FontItalic = False
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub Form_Load()
'Me.Height = 3375
Banner_IMAGE.Picture = LoadResPicture("POLIN2", vbResBitmap)
aplicar_transparencia Me.hwnd, 240
Me.Icon = WinSeek.Icon
icono.Picture = WinSeek.Icon
Me.Caption = LoadResString(43) & " " & LoadResString(1) & " " & LoadResString(2)
lbl_titleyversion.Caption = LoadResString(1) & " " & LoadResString(2)
'lbl_autor.Caption = LoadResString(3)
lbl_fecha.Caption = LoadResString(5)
lbl_url.Caption = LoadResString(4)

mostrar_systray_notificacion LoadResString(49), LoadResString(31)
Text1.Text = LoadResString(6)

Licence_User.Caption = LoadResString(68) & " " & get_Usuario
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_url.FontItalic = False
aplicar_transparencia Me.hwnd, 240
End Sub


Private Sub icono_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub icono_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_autor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_autor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_correo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_correo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_fecha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_fecha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_titleyversion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_titleyversion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aplicar_transparencia Me.hwnd, 240
End Sub

Private Sub lbl_url_Click()
Shell Environ("programfiles") & "/Internet Explorer/iexplore.exe " & lbl_url.Caption, vbMaximizedFocus
End Sub

Private Sub lbl_url_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_url.FontItalic = True
lbl_url.MousePointer = 99
End Sub

   
 'retorna un String con el nombre de usuario actual de windows
 '***************************************************************
 Private Function get_Usuario() As String
       
     Dim Nombre As String, Ret As Long
       
     ' Buffer
     Nombre = Space$(250)
       
     ' TamaÃ±o
     Ret = Len(Nombre)
       
     If GetUserName(Nombre, Ret) = 0 Then
         get_Usuario = vbNullString
     Else
         ' Extrae solo los caracteres
         get_Usuario = Left$(Nombre, Ret - 1)
     End If
       
 End Function

Private Sub Licence_User_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
aplicar_transparencia Me.hwnd, 100
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Licence_User_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_url.FontItalic = False
aplicar_transparencia Me.hwnd, 240
End Sub
