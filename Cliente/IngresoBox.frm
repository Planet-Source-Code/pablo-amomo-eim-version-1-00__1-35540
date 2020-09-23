VERSION 5.00
Begin VB.Form IngresoBox 
   BorderStyle     =   0  'None
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   7035
   Icon            =   "IngresoBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "IngresoBox.frx":000C
   ScaleHeight     =   1875
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Ingreso 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   840
      Width           =   6555
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   6600
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6705
      MouseIcon       =   "IngresoBox.frx":2576
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
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
      Height          =   270
      Left            =   160
      TabIndex        =   4
      Top             =   520
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label BotonCancelar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "IngresoBox.frx":26C8
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
   End
   Begin VB.Label BotonOk 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5190
      MouseIcon       =   "IngresoBox.frx":281A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1290
      Width           =   1605
   End
   Begin VB.Label LabelCancelar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1410
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
      Left            =   5190
      TabIndex        =   0
      Top             =   1410
      Width           =   1575
   End
   Begin VB.Shape ShapeCancelar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   5190
      Shape           =   4  'Rounded Rectangle
      Top             =   1305
      Width           =   1575
   End
   Begin VB.Shape ShapeOk 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   1305
      Width           =   1575
   End
End
Attribute VB_Name = "IngresoBox"
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
Public Modal As Boolean
Option Explicit
Public Sub CargarDatosInputBox(Texto As String, LargoIngreso As Integer, TituloVentana As String)

 ' **************************************************************
 ' Carga la Info correspondiente a la Ventan
 ' **************************************************************
 Me.TituloVentana1 = Trim(TituloVentana) & "..."
 Me.TextoMensaje = Texto
 Me.Ingreso.MaxLength = LargoIngreso
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 
End Sub
Private Sub BotonCancelar_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeCancelar

 Variables.RespuestaIngresoBox = ""
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

 Variables.RespuestaIngresoBox = Ingreso
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
 ' Carga los textos del Formularioel Icono de Aplicacion
 ' **************************************************************
 Me.TextoMensaje.ForeColor = Variables.FontLabelColor
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 Me.ShapeCancelar.BackColor = Variables.ShapesBackColor
 Me.ShapeCancelar.BorderColor = Variables.ShapesBorderColor
 Me.LabelCancelar.ForeColor = Variables.FontBotonesColor
 Me.LabelCancelar = MensajeRecurso(106)     ' Cancelar
 Me.ShapeOk.BackColor = Variables.ShapesBackColor
 Me.ShapeOk.BorderColor = Variables.ShapesBorderColor
 Me.LabelOkLBL.ForeColor = Variables.FontBotonesColor
 Me.LabelOkLBL = MensajeRecurso(275)        ' Ok

End Sub
Private Sub Form_Load()

 Me.FormularioNombre = "IngresoBox"
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.CargarTextos
 
 Me.Icon = Cliente.Icon
 
End Sub

Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Variables.RespuestaIngresoBox = ""
 Unload Me
 
End Sub

Private Sub Ingreso_KeyPress(KeyAscii As Integer)
 
 ' **************************************************************
 ' SI presiona ENTER simula como si ubiera clickeado
 ' en el Boton OK
 ' **************************************************************
 If KeyAscii = 13 Then
  BotonOk_Click
 End If
 
End Sub
