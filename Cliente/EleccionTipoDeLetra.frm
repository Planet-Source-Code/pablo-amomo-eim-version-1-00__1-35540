VERSION 5.00
Begin VB.Form EleccionTipoDeLetra 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   ClipControls    =   0   'False
   Icon            =   "EleccionTipoDeLetra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "EleccionTipoDeLetra.frx":000C
   ScaleHeight     =   2760
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4050
      TabIndex        =   9
      Top             =   1620
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2940
      TabIndex        =   8
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1620
      Width           =   2715
   End
   Begin VB.Shape LetraFondo 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   1350
      Top             =   480
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Letra 
      BackStyle       =   0  'Transparent
      Caption         =   "ABCabc"
      Height          =   255
      Index           =   0
      Left            =   150
      MouseIcon       =   "EleccionTipoDeLetra.frx":166A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   4740
      MouseIcon       =   "EleccionTipoDeLetra.frx":17BC
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2190
      Width           =   1545
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6270
      MouseIcon       =   "EleccionTipoDeLetra.frx":190E
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6225
   End
   Begin VB.Shape Shape20 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   315
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4710
      TabIndex        =   3
      Top             =   2325
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   270
      MouseIcon       =   "EleccionTipoDeLetra.frx":1A60
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2205
      Width           =   1575
   End
   Begin VB.Label CancelarLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   2325
      Width           =   1545
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   2185
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   4710
      Shape           =   4  'Rounded Rectangle
      Top             =   2185
      Width           =   1575
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Image FondoMedio 
      Height          =   420
      Left            =   0
      Picture         =   "EleccionTipoDeLetra.frx":1BB2
      Stretch         =   -1  'True
      Top             =   480
      Width           =   6585
   End
   Begin VB.Image FondoAbajo 
      Height          =   1260
      Left            =   0
      Picture         =   "EleccionTipoDeLetra.frx":AC54
      Top             =   1500
      Width           =   6585
   End
End
Attribute VB_Name = "EleccionTipoDeLetra"
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

Public FontActual As String
Public FormOrigen As Long
' **************************************************************

Option Explicit
Public CambiarUsuario As String
Public Sub MostrarFormulario(Fuente As String, FormularioOrigen As Long, Optional Modal As Boolean)

 Me.FontActual = Fuente
 ' Fuente:
 Me.Label4 = MensajeRecurso(272) & Fuente
 Me.Text1.Font = Fuente
 DistribuirFormulario
 ' Solo lo Muestra en Modal si es que se pide asi...
 If Modal Then
   Me.Show vbModal
  Else
   Me.Show
 End If
 FormOrigen = FormularioOrigen
  
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0
  Exit Sub
 End If

End Sub
Private Sub DistribuirFormulario()
Dim ContadorX, ContadorY, Contador As Integer
Dim PosicionX, PosicionY As Integer

 ' **************************************************************
 ' Pone las Letras
 ' **************************************************************
 ContadorX = 0
 ContadorY = 0
 For Contador = 1 To Variables.CantidadDeFuentes
  ContadorX = ContadorX + 1
  If ContadorX = 5 Then
   ContadorY = ContadorY + 1
   ContadorX = 1
  End If
  ' Carga el Label y el Fondo
  Load Letra((ContadorY * 4) + ContadorX)
  Load LetraFondo((ContadorY * 4) + ContadorX)
  ' Define el Tipo de Letra
  Letra(ContadorY * 4 + ContadorX).Font = Variables.NombreDeFuentes(Contador)
  Letra(ContadorY * 4 + ContadorX).FontSize = 8
  Letra(ContadorY * 4 + ContadorX).ForeColor = Variables.FontCambiarLetraLetra
  ' Pone en Posicion
  PosicionX = (ContadorX * 1550) - 1325
  PosicionY = 480 + (ContadorY * 270)
  Letra(ContadorY * 4 + ContadorX).Left = PosicionX
  Letra(ContadorY * 4 + ContadorX).Top = PosicionY
  LetraFondo(ContadorY * 4 + ContadorX).Left = PosicionX - 15
  LetraFondo(ContadorY * 4 + ContadorX).Top = PosicionY - 10
  ' Pone los Atributos
  Letra(ContadorY * 4 + ContadorX).Visible = True
  ' Si la Fuente es la actual la marca especial
  If UCase(FontActual) = UCase(NombreDeFuentes(Contador)) Then
    LetraFondo(ContadorY * 4 + ContadorX).Visible = True
    LetraFondo(ContadorY * 4 + ContadorX).BackColor = Variables.FontCambiarLetraNormalBack
    LetraFondo(ContadorY * 4 + ContadorX).BorderColor = Variables.FontCambiarLetraNormalBorder
   Else
    LetraFondo(ContadorY * 4 + ContadorX).Visible = False
    LetraFondo(ContadorY * 4 + ContadorX).BackColor = Variables.FontCambiarLetraHighBack
    LetraFondo(ContadorY * 4 + ContadorX).BorderColor = Variables.FontCambiarLetraHighBorder
  End If
  ' Los pone al frente
  LetraFondo(ContadorY * 4 + ContadorX).ZOrder (0)
  Letra(ContadorY * 4 + ContadorX).ZOrder (0)
  Letra(ContadorY * 4 + ContadorX).BackStyle = 0
  Letra(ContadorY * 4 + ContadorX).ToolTipText = Variables.NombreDeFuentes(Contador)
 Next
 
 ' **************************************************************
 ' Solo mueve los controles si es mas de 20
 ' **************************************************************
 If Variables.CantidadDeFuentes < 21 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Mueve los Controles como Corresponde
 ' **************************************************************
 ' **
 ' Fondo Medio
 ' **
 FondoMedio.HeighT = PosicionY
 
 ' **
 ' Fondo Abajo
 ' **
 FondoAbajo.Top = FondoMedio.Top + FondoMedio.HeighT
  ' **
 ' Tamaño del Form
 ' **
 Me.HeighT = FondoAbajo.Top + FondoAbajo.HeighT
 
 Dim DistanciaBot As Integer
 DistanciaBot = 690
 ' **
 ' Botone Cancelar
 ' **
 CancelarBt.BackColor = Variables.ShapesBackColor
 CancelarBt.BorderColor = Variables.ShapesBorderColor
 CancelarBt.Top = FondoAbajo.Top + DistanciaBot
 CancelarLbl.ForeColor = Variables.FontBotonesColor
 CancelarLbl.Top = FondoAbajo.Top + DistanciaBot + 110
 Label1.ForeColor = Variables.FontLabelColor
 Label1.Top = FondoAbajo.Top + DistanciaBot
 ' **
 ' Boton Ok
 ' **
 Label2.ForeColor = Variables.FontLabelColor
 Label2.Top = FondoAbajo.Top + DistanciaBot + 110
 Label3.Top = FondoAbajo.Top + DistanciaBot
 Shape3.BackColor = Variables.ShapesBackColor
 Shape3.BorderColor = Variables.ShapesBorderColor
 Shape3.Top = FondoAbajo.Top + DistanciaBot
 ' **
 ' Caja de Test
 ' **
 Label4.ForeColor = Variables.FontLabelColor
 Label4.Top = FondoAbajo.Top + 160
 Label5.ForeColor = Variables.FontLabelColor
 Label5.Top = FondoAbajo.Top + 160
 Text1.Top = FondoAbajo.Top + 130
 
End Sub
Public Sub CargarTextos()

 ' **************************************************************
 ' Carga los Textos del Formulario
 ' **************************************************************
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(272)        ' Nombre de la Fuente:
 Me.Label5.ForeColor = Variables.FontLabelColor
 Me.Label5 = MensajeRecurso(273)        ' Area de Test
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.CancelarLbl = MensajeRecurso(106)   ' Cancelar
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(275)        ' Ok
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 
End Sub

Private Sub Form_Load()
 
 ' **************************************************************
 ' Carga el Titulo de la Ventana
 ' **************************************************************
 '  - Eleccion De Fuente...
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(276)
 Me.CargarTextos
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Icon = Cliente.Icon
 
End Sub
Private Sub BuscarDescargar()
Dim Contador As Integer

 ' **************************************************************
 ' Busca el formulario
 ' **************************************************************
 For Contador = 1 To Forms.Count - 1
  If Forms(Contador).hwnd = Me.FormOrigen Then
   Forms(Contador).CambiarLetraRemoto (Me.FontActual)
   Exit For
  End If
 Next
  
 ' **************************************************************
 ' Descarga el Formulario
 ' **************************************************************
 Unload Me

End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Me.FontActual = ""
 
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
  
 ' **************************************************************
 ' Pasar el Font y Descargar
 ' **************************************************************
 BuscarDescargar
 
 
End Sub
Private Sub Label1_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.CancelarBt

 Me.FontActual = ""
 
 ' **************************************************************
 ' Pasar el Font y Descargar
 ' **************************************************************
 BuscarDescargar
 
 
End Sub

Private Sub Label3_Click()

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
  
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape3
  
 ' **************************************************************
 ' Pasar el Font y Descargar
 ' **************************************************************
 BuscarDescargar
 
End Sub
Private Sub Letra_Click(Index As Integer)

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
 
 ' **************************************************************
 ' Define la Letra Actual
 ' **************************************************************
 Me.FontActual = Letra(Index).ToolTipText
 ' Fuente:
 Me.Label4 = MensajeRecurso(272) & Me.FontActual
 Me.Text1.Font = Me.FontActual
 
End Sub
Private Sub Letra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Contador As Integer
Dim LetraActual As Integer

 ' **************************************************************
 ' Busca el Indice de la Letra Actual
 ' **************************************************************
 For Contador = 1 To Variables.CantidadDeFuentes
  If UCase(Letra(Contador).ToolTipText) = UCase(FontActual) Then
   LetraActual = Contador
  End If
 Next
  
 ' **************************************************************
 ' Elimina todos los Fondos visibles
 ' **************************************************************
 For Contador = 1 To Variables.CantidadDeFuentes
  ' Si es el Font Actual lo deja en blanco
  If Contador = LetraActual Then
    LetraFondo(Contador).Visible = True
    LetraFondo(Contador).BackColor = Variables.FontCambiarLetraNormalBack
    LetraFondo(Contador).BorderColor = Variables.FontCambiarLetraNormalBorder
   Else
    ' Esto es para evitar el Parpadeo cuando ya esta marcado...
    If Index <> Contador Then
     LetraFondo(Contador).Visible = False
    End If
  End If
 Next
 
 ' **************************************************************
 ' Setea el Fondo donde esta el Mouse
 ' **************************************************************
 LetraFondo(Index).BackColor = Variables.FontCambiarLetraHighBack
 LetraFondo(Index).BorderColor = Variables.FontCambiarLetraHighBorder
 LetraFondo(Index).Visible = True
 
End Sub
