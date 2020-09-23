VERSION 5.00
Begin VB.Form AgregarBuscarAmigo 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   ClipControls    =   0   'False
   Icon            =   "AgregarBuscarAmigo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AgregarBuscarAmigo.frx":000C
   ScaleHeight     =   4260
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   180
      ScaleHeight     =   90
      ScaleWidth      =   4995
      TabIndex        =   17
      Top             =   3330
      Width           =   4995
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   150
      ScaleHeight     =   60
      ScaleWidth      =   4995
      TabIndex        =   16
      Top             =   1740
      Width           =   4995
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   150
      ScaleHeight     =   1770
      ScaleWidth      =   45
      TabIndex        =   15
      Top             =   1680
      Width           =   45
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   5010
      ScaleHeight     =   1725
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Grupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   13
      Top             =   820
      Width           =   1845
   End
   Begin VB.Timer Animacion 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4590
      Top             =   660
   End
   Begin VB.TextBox AmigoBuscarApellidoYNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1150
      Width           =   3015
   End
   Begin VB.TextBox Amigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   0
      Top             =   490
      Width           =   2955
   End
   Begin VB.ListBox ResultadoBusqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1590
      ItemData        =   "AgregarBuscarAmigo.frx":330A
      Left            =   150
      List            =   "AgregarBuscarAmigo.frx":330C
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   1770
      Width           =   5115
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3900
      MouseIcon       =   "AgregarBuscarAmigo.frx":330E
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   240
   End
   Begin VB.Image ScrollAbajo 
      Height          =   240
      Left            =   5340
      MouseIcon       =   "AgregarBuscarAmigo.frx":3460
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image ScrollArriba 
      Height          =   240
      Left            =   5340
      MouseIcon       =   "AgregarBuscarAmigo.frx":35B2
      MousePointer    =   99  'Custom
      Top             =   1650
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   5410
      MouseIcon       =   "AgregarBuscarAmigo.frx":3704
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   -30
      TabIndex        =   8
      Top             =   30
      Width           =   5025
   End
   Begin VB.Label Label7 
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
      Height          =   225
      Left            =   210
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   5340
      Top             =   450
      Width           =   270
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2070
      MouseIcon       =   "AgregarBuscarAmigo.frx":3856
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3675
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Height          =   225
      Left            =   210
      TabIndex        =   9
      Top             =   1170
      Width           =   1965
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
      TabIndex        =   7
      Top             =   120
      Width           =   4245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   3900
      MouseIcon       =   "AgregarBuscarAmigo.frx":39A8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3900
      TabIndex        =   6
      Top             =   3815
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "AgregarBuscarAmigo.frx":3AFA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3690
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
      TabIndex        =   5
      Top             =   3815
      Width           =   1545
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label UsuarioLbl 
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
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   510
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   3900
      Shape           =   4  'Rounded Rectangle
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2085
      TabIndex        =   11
      Top             =   3815
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2085
      Shape           =   4  'Rounded Rectangle
      Top             =   3690
      Width           =   1575
   End
End
Attribute VB_Name = "AgregarBuscarAmigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **************************************************************
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
' **************************************************************

' **************************************************************
' Define los Menus...
' **************************************************************
Public WithEvents MenuDeGrupos As IcoMenu
Attribute MenuDeGrupos.VB_VarHelpID = -1
Private IndiceAnimacion As Integer
' **************************************************************

' **************************************************************
' Define el Grupo al cual se desea enviar el Amigo, lo utiliza
' el Menu Descolgable que permite Mover Amigos...
' **************************************************************
Public NombreDelGrupo As String
' **************************************************************
Private Sub Amigo_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Enter cumple la funcion como si se apretara Buscar...
 ' **************************************************************
 If KeyAscii = 13 Then
  Label5_Click
 End If
  
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Funcion que permite realizar el Move del Formulario...
 ' **************************************************************
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0
  Exit Sub
 End If

End Sub
Public Sub CargarTextos()

 ' **************************************************************
 ' Define los Textos del Formulario...
 ' **************************************************************
 '  Titulo...
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(131)
 ' Lables...
 Me.UsuarioLbl.ForeColor = Variables.FontLabelColor
 Me.UsuarioLbl = MensajeRecurso(125) ' Alias del Amigo:
 Me.Label7.ForeColor = Variables.FontLabelColor
 Me.Label7 = MensajeRecurso(126) ' Grupo Seleccionado:
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(451) ' Cadena a Buscar:
 ' Botones...
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.CancelarLbl = MensajeRecurso(128) ' Cerrar...
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.Label6.ForeColor = Variables.FontBotonesColor
 Me.Label6 = MensajeRecurso(129) ' Buscar Amigo...
 Me.Shape6.BackColor = Variables.ShapesBackColor
 Me.Shape6.BorderColor = Variables.ShapesBorderColor
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(130) ' Agregar Amigo...
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 Me.Image2.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.ScrollArriba.Picture = Cliente.ImagenesFlecha.ListImages("ArribaAzul").Picture
 Me.ScrollAbajo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoAzul").Picture
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 
End Sub
Private Sub Form_Load()
Dim Contador As Integer
 
 ' **************************************************************
 ' Carga el Titulo de la Ventana
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
 ' **************************************************************
 ' Determina el Formulario Tipo a traves del Nombre
 ' **************************************************************
 FormularioNombre = "AgregarBuscarAmigos"
 
 ' **************************************************************
 ' Menu De Cambio de Grupo
 ' **************************************************************
  Set Me.MenuDeGrupos = New IcoMenu
  With Me.MenuDeGrupos
   ' Fuera de Grupo...
   .SetItem 0, MensajeRecurso(132), Cliente.ImagenesAmigos.ListImages("Grupo").Picture, MensajeRecurso(132)
   ' **************************************************************
   ' Carga Todos los Grupos en el Listado Despleglable...
   ' **************************************************************
   Dim Cantidad As Integer
   Cantidad = 0
   For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
    If Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1) = "G" Then
     Cantidad = Cantidad + 1
     .SetItem 0 + Cantidad, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2), Cliente.ImagenesAmigos.ListImages("Grupo").Picture, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2)
    End If
   Next
  End With
  ' Por default Muestra el Item Fuera de Grupo...
  Me.Grupo = MensajeRecurso(132)
  
End Sub
Private Sub Image2_Click() ' Muestra el Menu desplegable de Grupo...

 ' **************************************************************
 ' Ejecuta el Sonido de Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Me.MenuDeGrupos.ShowMenu Me.Image2.Left + Me.Left, Me.Image2.Top + Me.Top + 230

End Sub
Private Sub MenuDeGrupos_Click(ByVal Index As Long, Tag As String)

 ' **************************************************************
 ' Define el Grupo Seleccionado en el Menu Desplegable...
 ' NOTA: En el TAG del Menu se carga el Nombre del Mismo...
 ' **************************************************************
 Me.Grupo = Tag
 
End Sub
Private Sub Image3_Click()
 
 ' **************************************************************
 ' Ejecuta el Sonido de Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Descarga el Formulario...
 ' **************************************************************
 Unload AgregarBuscarAmigo
 
End Sub
Private Sub Label1_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton...
 ' **************************************************************
 EfectoBoton Me.CancelarBt
 
 ' **************************************************************
 ' Descarga el Formulario...
 ' **************************************************************
 Unload AgregarBuscarAmigo
 
End Sub
Private Sub Label3_Click()
Dim AmigoACrear, NombreGrupo, EstadoAmigoTexto, Sexo As String
Dim Respuesta, Contador, Estadoamigo As Integer
'Dim SegundosTranscurridos As Integer
'Dim TiempoInicial As Date
Dim Usuarioexiste As Boolean

 ' **************************************************************
 ' Ejecuta el Sonido de Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton...
 ' **************************************************************
 EfectoBoton Me.Shape3

 ' **************************************************************
 ' Verifica en que grupo debe crear al Amigo...
 ' **************************************************************
 ' Fuera de Grupo...
 If IsNull(Me.Grupo) Or Me.Grupo = MensajeRecurso(132) Then
   NombreGrupo = ""
  Else
   NombreGrupo = Trim(Me.Grupo)
 End If
 
 ' **************************************************************
 ' Levanta el Nombre del Amigo a Agregar...
 ' **************************************************************
 AmigoACrear = Trim(Me.Amigo)
 
 ' **************************************************************
 ' Verifica el Amigo a crear... (No puede ser Nulo)
 ' **************************************************************
 If AmigoACrear = "" Then
  ' Muestra: Ingrese un Nombre de Amigo Valido...
  Respuesta = Varios.MostrarMSGBox(MensajeRecurso(133), vbOKOnly, "vbCritical", Configuracion.TituloVentanas)
  Exit Sub
 End If
 ' **************************************************************
 ' Verificar si el Nuevo Amigo a Agregar no Existe...
 ' **************************************************************
 If EsAmigo(CStr(AmigoACrear)) Then
 'For Contador = 1 To Variables.CantidadGrupoAmigo
 ' If UCase(Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(AmigoACrear) Then
   ' Muestra: El Amigo [ % ] ya Existe en su Listado de Amigos...
   MostrarMSGBox MensajeRecurso(134) & AmigoACrear & MensajeRecurso(135), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
 '  Exit Sub
 ' End If
 'Next
 End If
 ' **************************************************************
 ' Desactiva los Botones y Activa la Animacion (De Tasmision)...
 ' **************************************************************
 ActivaDesactivarBotones False
 'Animacion.Enabled = True
 'Me.Label1.Enabled = False
 'Me.Label5.Enabled = False
 'Me.Label3.Enabled = False
 'Me.Amigo.Enabled = False
 'Me.AmigoBuscarApellidoYNombre.Enabled = False
 'Me.Grupo.Enabled = False
 'Me.ResultadoBusqueda.Enabled = False
 
 Respuesta = SolicitarAgregarAmigo(CStr(AmigoACrear))
 
 ' **************************************************************
 ' Desactiva los Botones y Activa la Animacion
 ' **************************************************************
 ActivaDesactivarBotones True
 'Animacion.Enabled = False
 'Me.Label1.Enabled = True
 'Me.Label5.Enabled = True
 'Me.Label3.Enabled = True
 'Me.Amigo.Enabled = True
 'Me.AmigoBuscarApellidoYNombre.Enabled = True
 'Me.Grupo.Enabled = True
 'Me.ResultadoBusqueda.Enabled = True
 'AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 
 
 ' **************************************************************
 ' Verifica la Respuesta
 ' **************************************************************
 Estadoamigo = 0
 EstadoAmigoTexto = ""
 Select Case Respuesta
  Case -1
   ' Muestra: No se Consiguio Respuesta del Servidor, ¿Desea Agregar al Amigo como 'Usuario Inexistente'?
   Respuesta = Varios.MostrarMSGBox(MensajeRecurso(136), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then
    Exit Sub
   End If
   Usuarioexiste = False
  Case 0
   ' Muestra: El Amigo [ % ] no existe... ¿Desea Agregar al Amigo como 'Usuario Inexistente'?
   Respuesta = Varios.MostrarMSGBox(MensajeRecurso(134) & Trim(AmigoACrear) & MensajeRecurso(137), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then
    Exit Sub
   End If
   Usuarioexiste = False
  Case 1
   Estadoamigo = Variables.RespuestaAmigoACrearEstado
   EstadoAmigoTexto = Variables.RespuestaAmigoACrearEstadoTexto
   Sexo = Variables.RespuestaAmigoACrearSexo
   Usuarioexiste = True
   ' Muestra: Ok no hace nada y Agrega el Amigo, en estado desconectado...
 End Select
 
 ' **************************************************************
 ' Crea el Nuevo Amigo
 ' **************************************************************
 Varios.CrearNuevoAmigo Usuarioexiste, CStr(NombreGrupo), CStr(AmigoACrear), CInt(Estadoamigo), CStr(EstadoAmigoTexto), CStr(UCase(Sexo))
 
 ' **************************************************************
 ' Blanquea el Campo de Creación de Amigo...
 ' **************************************************************
 AgregarBuscarAmigo.Amigo = ""
 
End Sub
Private Sub Animacion_Timer()

 ' **************************************************************
 ' Timer que controla la animacion de la Conección
 ' **************************************************************
 ' Si el timer no es True sale...
 If Animacion = False Then Exit Sub
 ' Verifica que figura debe mostrar...
 IndiceAnimacion = IndiceAnimacion + 1
 If IndiceAnimacion = 5 Then IndiceAnimacion = 2
 ' Pone la Imagen...
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(IndiceAnimacion).Picture
  
End Sub
Private Sub Label5_Click() ' Buscar Amigo...
Dim Contador, SegundosTranscurridos, Respuesta As Integer
Dim TiempoInicial As Date
Dim CadenaBusqueda As String

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape6
 
 ' **************************************************************
 ' Define la Cadena de Busqueda...
 ' **************************************************************
 CadenaBusqueda = Trim(AgregarBuscarAmigo.AmigoBuscarApellidoYNombre)
 
 ' **************************************************************
 ' Enviar la Query al Servidor Sobre el Amigo...
 ' **************************************************************
 EnviarPaqueteTCP ("25" & CompletarCadena(CStr(CadenaBusqueda), 50, "D", " "))
 
 ' **************************************************************
 ' Desactiva los Botones y Activa la Animación...
 ' **************************************************************
 ActivaDesactivarBotones False
 'Animacion.Enabled = True
 'Me.Label1.Enabled = False
 'Me.Label5.Enabled = False
 'Me.Label3.Enabled = False
 'Me.Amigo.Enabled = False
 'Me.Grupo.Enabled = False
 'Me.AmigoBuscarApellidoYNombre.Enabled = False
 'Me.ResultadoBusqueda.Enabled = False
 
 ' **************************************************************
 ' Espera 5 segundos por el OK de la Busqueda...
 ' **************************************************************
 Variables.RespuestaBusquedaAmigos = -1
 TiempoInicial = Time
 Do
  DoEvents
  If Variables.RespuestaBusquedaAmigos <> -1 Then Exit Do
  'SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
  'If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
  ' Tardo mas de 5 segundos... (SALE !)
  If DateDiff("s", TiempoInicial, Time) >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 
 ' **************************************************************
 ' Activa los Botones y DesActiva la Animacion...
 ' **************************************************************
 ActivaDesactivarBotones True
 'Animacion.Enabled = False
 'Me.Label1.Enabled = True
 'Me.Label5.Enabled = True
 'Me.Label3.Enabled = True
 'Me.Amigo.Enabled = True
 'Me.AmigoBuscarApellidoYNombre.Enabled = True
 'Me.Grupo.Enabled = True
 'Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 'Me.ResultadoBusqueda.Enabled = True
 
 ' **************************************************************
 ' Verifica la Respuesta
 ' **************************************************************
 Select Case Variables.RespuestaBusquedaAmigos
  Case -1
   ' Muestra: No se Recibio Respuesta del Servidor para la Búsqueda de Amigos...
   Respuesta = Varios.MostrarMSGBox(MensajeRecurso(138), vbOKOnly, "vbCritical", Configuracion.TituloVentanas)
  Case 0
   ' Muestra: No se Encontraron Amigos con la Cadena [ % ]...
   Respuesta = Varios.MostrarMSGBox(MensajeRecurso(139) & Trim(Me.AmigoBuscarApellidoYNombre) & MensajeRecurso(121), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
  Case 1
   ' Muestra: Ok no hace nada...
 End Select

End Sub
Private Sub ResultadoBusqueda_DblClick()
Dim Posicion, Respuesta As Integer
Dim Cadena, Usuario As String

 ' **************************************************************
 ' Cuando se clickea hace un Agregar Amigo
 ' **************************************************************
 ' Verifica el Usuario que se quiere agregar
 Cadena = ResultadoBusqueda.List(Me.ResultadoBusqueda.ListIndex)
 Posicion = InStr(1, Cadena, "(", vbTextCompare)
 Usuario = Trim(Mid$(Cadena, 2, Posicion - 5))
 
 ' **************************************************************
 ' Pregunta si quiere agregar el Usuario
 ' **************************************************************
 ' Muestra: ¿Desea Agregar el Usuario [ % ] al Grupo [ % ] en su Listado de Amigos?
 Respuesta = Varios.MostrarMSGBox(MensajeRecurso(141) & Usuario & MensajeRecurso(142) & Trim(Me.Grupo) & MensajeRecurso(143), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
 If Respuesta = vbNo Then
  Exit Sub
 End If
    
 ' **************************************************************
 ' Agrega el Usuario
 ' **************************************************************
 Me.Amigo = Usuario
 Label3_Click
 
End Sub
Private Sub ScrollAbajo_Click()

 ' **************************************************************
 ' Hace el Scroll del ListBox... (DOWN)
 ' **************************************************************
 Scroller.ScrollListBox Me.ResultadoBusqueda, "Abajo"

End Sub
Private Sub ScrollArriba_Click()

 ' **************************************************************
 ' Hace el Scroll del ListBox... (UP)
 ' **************************************************************
 Scroller.ScrollListBox Me.ResultadoBusqueda, "Arriba"
 
End Sub
Private Sub ActivaDesactivarBotones(Estado As Boolean)

 ' **************************************************************
 ' Desactiva los Botones y Activa la Animacion (De Tasmision)...
 ' **************************************************************
 Animacion.Enabled = Not Estado
 Me.Label1.Enabled = Estado
 Me.Label5.Enabled = Estado
 Me.Label3.Enabled = Estado
 Me.Amigo.Enabled = Estado
 Me.AmigoBuscarApellidoYNombre.Enabled = Estado
 Me.Grupo.Enabled = Estado
 Me.ResultadoBusqueda.Enabled = Estado
 ' Si se activan los Botones pone la imagen de Coneccion apagado...
 If Estado = True Then
  AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 End If
 
End Sub
