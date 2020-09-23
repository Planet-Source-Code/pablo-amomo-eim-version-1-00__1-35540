VERSION 5.00
Begin VB.Form MensajesBox 
   BorderStyle     =   0  'None
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   7035
   Icon            =   "MensajesBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "MensajesBox.frx":000C
   ScaleHeight     =   1710
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6720
      MouseIcon       =   "MensajesBox.frx":24EF
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6645
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   3315
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Label TextoMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   750
      TabIndex        =   6
      Top             =   450
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImagenMensaje 
      Height          =   480
      Left            =   180
      Top             =   460
      Width           =   480
   End
   Begin VB.Label BotonNo 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   4455
      MouseIcon       =   "MensajesBox.frx":2641
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label BotonOk 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2715
      MouseIcon       =   "MensajesBox.frx":2793
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label BotonSi 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   975
      MouseIcon       =   "MensajesBox.frx":28E5
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label LabelNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4470
      TabIndex        =   2
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label LabelOkLBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2730
      TabIndex        =   1
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label LabelSi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   975
      TabIndex        =   0
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Shape ShapeNo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Shape ShapeOk 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2730
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Shape ShapeSi 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   975
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   1575
   End
End
Attribute VB_Name = "MensajesBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **************************************************************
' Variables Compartidas de Formulario (Para Busqueda)
' **************************************************************
Public FormularioNombre As String
Public AliasUsuario, LabelOk As String
Public Refresco As Boolean
Public GraboLosDatosUsuario As Boolean
Public CambioDeDatosUsuario As Boolean
Public Grabando, Refrescando As Boolean
' **************************************************************

' **************************************************************
Public BotonesFormulario As Long
Public Modal As Boolean
Option Explicit
Public Sub CargarDatosBox(Texto As String, Botones As Integer, Imagen As String, TituloVentana As String)

 ' **************************************************************
 ' Carga la Info correspondiente a la Ventan
 ' **************************************************************
 Me.TituloVentana1 = Trim(TituloVentana) & "..."
 Me.TextoMensaje = Texto
 Me.ImagenMensaje.Picture = Cliente.ListadoImagenes.ListImages(Imagen).Picture
 
 ' **************************************************************
 ' Elige y define los Botones a Mostrar
 ' **************************************************************
 BotonesFormulario = Botones
 Select Case Botones
  Case vbOKOnly:
   ShapeSi.Visible = False
   BotonSi.Enabled = False
   LabelSi.Visible = False
   ShapeNo.Visible = False
   BotonNo.Enabled = False
   LabelNo.Visible = False
   ShapeOk.Visible = True
   BotonOk.Enabled = True
   LabelOkLBL.Visible = True
  Case vbYesNo:
   ShapeSi.Visible = True
   BotonSi.Enabled = True
   LabelSi.Visible = True
   ShapeNo.Visible = True
   BotonNo.Enabled = True
   LabelNo.Visible = True
   ShapeOk.Visible = False
   BotonOk.Enabled = False
   LabelOkLBL.Visible = False
 End Select

End Sub
Private Sub BotonNo_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeNo

 Variables.RespuestaMensajeBox = vbNo
 Unload Me
 
End Sub
Private Sub BotonOk_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeOk

 Variables.RespuestaMensajeBox = vbOK
 Unload Me
 
End Sub
Private Sub BotonSi_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeSi

 Variables.RespuestaMensajeBox = vbYes
 Unload Me
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 ReleaseCapture
 SendMessage Me.hwnd, &HA1, 2, 0
 Exit Sub
End If

End Sub
Public Sub CargarTextos()

 ' **************************************************************
 ' Carga los textos del Formulario
 ' **************************************************************
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 Me.TextoMensaje.ForeColor = Variables.FontLabelColor
 ' Carga los Labels
 Me.ShapeOk.BackColor = Variables.ShapesBackColor
 Me.ShapeOk.BorderColor = Variables.ShapesBorderColor
 Me.LabelOkLBL.ForeColor = Variables.FontBotonesColor
 Me.LabelOkLBL = MensajeRecurso(275)    ' Ok
 Me.ShapeSi.BackColor = Variables.ShapesBackColor
 Me.ShapeSi.BorderColor = Variables.ShapesBorderColor
 Me.LabelSi.ForeColor = Variables.FontBotonesColor
 Me.LabelSi = MensajeRecurso(289)       ' Si
 Me.ShapeNo.BackColor = Variables.ShapesBackColor
 Me.ShapeNo.BorderColor = Variables.ShapesBorderColor
 Me.LabelNo.ForeColor = Variables.FontBotonesColor
 Me.LabelNo = MensajeRecurso(288)       ' No
 ' Carga Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 
 If UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(275), 1, 1)) Or UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(289), 1, 1)) Or UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(288), 1, 1)) Then
  Select Case BotonesFormulario
   Case vbOKOnly:
    If UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(275), 1, 1)) Then BotonOk_Click
   Case vbYesNo:
    If UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(289), 1, 1)) Then BotonSi_Click
    If UCase(Chr(KeyAscii)) = UCase(Mid$(MensajeRecurso(288), 1, 1)) Then BotonNo_Click
  End Select
 End If
 
 If KeyAscii = 13 Or KeyAscii = 32 Then
  Select Case BotonesFormulario
   Case vbOKOnly:
    BotonOk_Click
   Case vbYesNo:
    BotonSi_Click
  End Select
 End If
 
End Sub
Private Sub Form_Load()

 ' **************************************************************
 ' Nombre del Formulario...
 ' **************************************************************
 Me.FormularioNombre = "MensajesBox"

 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Simula como Si determinara un No... o un OK...
 ' **************************************************************
 Select Case BotonesFormulario
  Case vbOKOnly:
   BotonOk_Click
  Case vbYesNo:
   BotonNo_Click
 End Select

End Sub
