VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MensajeOffLine 
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   Icon            =   "MensajeOffLine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MensajeOffLine.frx":000C
   ScaleHeight     =   2700
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Animacion 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   660
      Top             =   1590
   End
   Begin VB.TextBox MensajeOfflinePara 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox VentanaMensajes 
      Height          =   1095
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"MensajeOffLine.frx":2283
      MouseIcon       =   "MensajeOffLine.frx":22FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   2047
      Top             =   2250
      Width           =   270
   End
   Begin VB.Label BotonEnviar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "MensajeOffLine.frx":245C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Label BotonCancelar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   270
      MouseIcon       =   "MensajeOffLine.frx":25AE
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Label LabelEnviarMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2535
      TabIndex        =   7
      Top             =   2245
      Width           =   1560
   End
   Begin VB.Label LabelCancelar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   255
      TabIndex        =   6
      Top             =   2245
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2535
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4065
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3315
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "MensajeOffLine.frx":2700
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Image ScrollAbajo 
      Height          =   240
      Left            =   3975
      MouseIcon       =   "MensajeOffLine.frx":2852
      MousePointer    =   99  'Custom
      Top             =   1340
      Width           =   240
   End
   Begin VB.Image ScrollArriba 
      Height          =   240
      Left            =   3975
      MouseIcon       =   "MensajeOffLine.frx":29A4
      MousePointer    =   99  'Custom
      Top             =   460
      Width           =   240
   End
End
Attribute VB_Name = "MensajeOffLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' **************************************************************
' Variables Compartidas de Formulario (Para Busqueda)
' **************************************************************
Public FormularioNombre As String
Public AliasUsuario, LabelOk As String
Public Refresco As Boolean
Public GraboLosDatosUsuario As Boolean
Public CambioDeDatosUsuario As Boolean
Public Grabando, Refrescando As Boolean
Public RecibidoOk As Integer

Private IndiceAnimacion As Integer
' **************************************************************
Private Function EnviarMensajeOffLine(Para As String, Mensaje As String) As Integer
Dim TiempoLogueoInicial As Date
    
  ' **************************************************************
  ' Enviar el Mensaje OffLine...
  ' **************************************************************
  EnviarPaqueteTCP "5" & CompletarCadena(CStr(Para), 16, "D", " ") & Mensaje

  ' **************************************************************
  ' Espera la Respuesta...
  ' **************************************************************
  TiempoLogueoInicial = Time
  RecibidoOk = -1
  Do Until RecibidoOk <> -1
   DoEvents
   If DateDiff("s", TiempoLogueoInicial, Time) >= Configuracion.TimeOutGeneral Then Exit Do
  Loop

  EnviarMensajeOffLine = RecibidoOk
  
End Function

Private Sub Animacion_Timer()
 ' **************************************************************
 ' Timer que controla la animacion de la Conección
 ' **************************************************************
 ' Si el timer no es True sale
 If Animacion = False Then Exit Sub
 
 ' Verifica que figura debe mostrar
 IndiceAnimacion = IndiceAnimacion + 1
 If IndiceAnimacion = 5 Then IndiceAnimacion = 2
 ' Muestra la imagen Correspondiente...
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(IndiceAnimacion).Picture

End Sub

Private Sub BotonCancelar_Click()

  ' **************************************************************
  ' Ejecuta el Sonido de Click
  ' **************************************************************
  Audio.EjecutarSonido "003"

  ' **************************************************************
  ' Descarga el Formulario...
  ' **************************************************************
  Unload Me
  
End Sub
Private Sub BotonEnviar_Click()
Dim Respuesta As Integer

  
  ' **************************************************************
  ' Ejecuta el Sonido de Click
  ' **************************************************************
  Audio.EjecutarSonido "003"
   
  ' **************************************************************
  ' Verificar que el Mensaje no sea demasiado grande
  ' **************************************************************
  If Len(Me.VentanaMensajes.TextRTF) > 4000 Then
   MostrarMSGBox MensajeRecurso(391), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Exit Sub
  End If
   
  ' **************************************************************
  ' Si no hay Nada que enviar sale...
  ' **************************************************************
  If Trim(Me.VentanaMensajes.Text) = "" Then Exit Sub
  
  ' **************************************************************
  ' Dispara el Timer de La Animacion (Comunicacion)
  ' **************************************************************
  Animacion.Enabled = True
  
  ' **************************************************************
  ' Lockea la Ventana de Mensajes...
  ' **************************************************************
  Me.VentanaMensajes.Locked = True
  Me.LabelCancelar.Enabled = False
  Me.LabelEnviarMensaje.Enabled = False
   
  ' **************************************************************
  ' Envia el Mensaje...
  ' **************************************************************
  Respuesta = EnviarMensajeOffLine(Me.AliasUsuario, VentanaMensajes.TextRTF)
  
  ' **************************************************************
  ' Detiene la Animacion...
  ' **************************************************************
  Animacion.Enabled = False
  AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 
  ' **************************************************************
  ' Libera la Ventana de Mensajes...
  ' **************************************************************
  Me.VentanaMensajes.Locked = False
  Me.LabelCancelar.Enabled = True
  Me.LabelEnviarMensaje.Enabled = True
  
  ' **************************************************************
  ' Procesa la Respuesta...
  ' **************************************************************
  'MsgBox RecibidoOk ok 456-457 / Error 458
  Select Case RecibidoOk
   Case 1:
    MostrarMSGBox MensajeRecurso(456) & "[" & Me.AliasUsuario & "]" & MensajeRecurso(457), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    Unload Me
   Case Else:
    MostrarMSGBox MensajeRecurso(458) & " [" & Me.AliasUsuario & "]...", vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  End Select

  
End Sub

Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Hace el Drag And Drop del Formulario...
 ' **************************************************************
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0
  Exit Sub
 End If
 
End Sub

Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos
 ' **************************************************************
 CargarTextos
 Me.Caption = Trim(Configuracion.TituloVentanas) & " - " & MensajeRecurso(455)
 Me.FormularioNombre = "MensajeOffline"
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Icon = Cliente.Icon

 ' **************************************************************
 ' Habilita la Ventana de Mensajes...
 ' **************************************************************
 Me.VentanaMensajes.Locked = False
 
End Sub
Private Sub CargarTextos()

 ' Titulo Ventana...
 TituloVentana1.ForeColor = Variables.FontTituloVentana
  TituloVentana1 = Trim(Configuracion.TituloVentanas) & " - " & MensajeRecurso(455)
 ' Botones...
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.LabelEnviarMensaje.ForeColor = Variables.FontBotonesColor
 Me.LabelEnviarMensaje = MensajeRecurso(218)    ' Enviar Mensaje...
 Me.Shape6.BackColor = Variables.ShapesBackColor
 Me.Shape6.BorderColor = Variables.ShapesBorderColor
 Me.LabelCancelar.ForeColor = Variables.FontBotonesColor
 Me.LabelCancelar = MensajeRecurso(106)  ' Cancelar
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.ScrollArriba.Picture = Cliente.ImagenesFlecha.ListImages("ArribaAzul").Picture
 Me.ScrollAbajo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoAzul").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture

End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Unload Me

End Sub

Private Sub ScrollAbajo_Click()
Dim Respuesta As Variant
    
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
    
 Respuesta = ScrollText&(VentanaMensajes, 1)

End Sub

Private Sub ScrollArriba_Click()
Dim Respuesta As Variant
    
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
    
 Respuesta = ScrollText&(VentanaMensajes, -1)

End Sub
