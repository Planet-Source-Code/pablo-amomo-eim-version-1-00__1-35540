VERSION 5.00
Begin VB.Form DatosUsuario 
   BorderStyle     =   0  'None
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   7035
   Icon            =   "DatosUsuario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "DatosUsuario.frx":000C
   ScaleHeight     =   6450
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox EstadoCivil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      Text            =   "EstadoCivil"
      Top             =   4590
      Width           =   1185
   End
   Begin VB.TextBox Signo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "Signo"
      Top             =   4260
      Width           =   1215
   End
   Begin VB.TextBox Sexo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "Sexo"
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Timer GrabacionTimeOut 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6510
      Top             =   4290
   End
   Begin VB.Timer Animacion 
      Interval        =   300
      Left            =   5580
      Top             =   4290
   End
   Begin VB.Timer TimeOut 
      Interval        =   500
      Left            =   6060
      Top             =   4290
   End
   Begin VB.TextBox OtraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "DatosUsuario.frx":4587
      Top             =   5235
      Width           =   4545
   End
   Begin VB.TextBox Telefono 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Text            =   "Telefono"
      Top             =   4935
      Width           =   3315
   End
   Begin VB.TextBox Ocupacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      Text            =   "Ocupacion"
      Top             =   3915
      Width           =   3315
   End
   Begin VB.TextBox Humor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      Text            =   "Humor"
      Top             =   3585
      Width           =   3315
   End
   Begin VB.TextBox FechaDeNacimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "FechaDeNac"
      Top             =   1920
      Width           =   4305
   End
   Begin VB.TextBox Intencion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "Intencion"
      Top             =   3255
      Width           =   3315
   End
   Begin VB.TextBox UbicacionGeografica 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "UbicacionGeografica"
      Top             =   2925
      Width           =   3315
   End
   Begin VB.TextBox Edad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "Edad"
      Top             =   2250
      Width           =   675
   End
   Begin VB.TextBox DireccionDeEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "DireccionDeEmail"
      Top             =   1590
      Width           =   4305
   End
   Begin VB.TextBox ApellidoYNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "ApellidoYNombre"
      Top             =   1260
      Width           =   4305
   End
   Begin VB.TextBox IDAliasUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "IDAliasUsuario"
      Top             =   930
      Width           =   2895
   End
   Begin VB.Image ComboEstadoCivil 
      Enabled         =   0   'False
      Height          =   240
      Left            =   3315
      MouseIcon       =   "DatosUsuario.frx":4592
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   240
   End
   Begin VB.Image ComboSigno 
      Enabled         =   0   'False
      Height          =   240
      Left            =   3315
      MouseIcon       =   "DatosUsuario.frx":46E4
      MousePointer    =   99  'Custom
      Top             =   4230
      Width           =   240
   End
   Begin VB.Image ComboSexo 
      Enabled         =   0   'False
      Height          =   240
      Left            =   3315
      MouseIcon       =   "DatosUsuario.frx":4836
      MousePointer    =   99  'Custom
      Top             =   2550
      Width           =   240
   End
   Begin VB.Image EstadoUsuarioImagen 
      Height          =   240
      Left            =   150
      Top             =   450
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   6645
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   36
      Top             =   120
      Width           =   6105
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6720
      MouseIcon       =   "DatosUsuario.frx":4988
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label BotonGrabar 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   435
      Left            =   2715
      MouseIcon       =   "DatosUsuario.frx":4ADA
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   5880
      Width           =   1605
   End
   Begin VB.Label LabelGrabar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3450
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape ShapeGrabar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2715
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label EstadoUsuarioTexto 
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
      Left            =   450
      TabIndex        =   32
      Top             =   480
      Width           =   6375
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   6630
      Top             =   450
      Width           =   270
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   31
      Top             =   5250
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   30
      Top             =   4950
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   29
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   28
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   27
      Top             =   3930
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   26
      Top             =   3630
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   25
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   24
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   23
      Top             =   2610
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   22
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   21
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   20
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   19
      Top             =   1290
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   210
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label BotonRefrescar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5160
      MouseIcon       =   "DatosUsuario.frx":4C2C
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5880
      Width           =   1605
   End
   Begin VB.Label BotonOk 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "DatosUsuario.frx":4D7E
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5880
      Width           =   1605
   End
   Begin VB.Label LabelRefrecar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5925
      TabIndex        =   15
      Top             =   6000
      Width           =   75
   End
   Begin VB.Label LabelOkLBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   255
      TabIndex        =   14
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Shape ShapeRefrescar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   5175
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Shape ShapeOk 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1575
   End
End
Attribute VB_Name = "DatosUsuario"
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

Option Explicit
Private CronometroInicio As Date
Private IndiceAnimacion As Integer
Public WithEvents MenuSexo As IcoMenu
Attribute MenuSexo.VB_VarHelpID = -1
Public WithEvents MenuEstadoCivil As IcoMenu
Attribute MenuEstadoCivil.VB_VarHelpID = -1
Public WithEvents MenuSigno As IcoMenu
Attribute MenuSigno.VB_VarHelpID = -1
Public EstadoNumero As String
Private Sub ComboEstadoCivil_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 CargaLosMenus "EstadoCivil"
 Me.MenuEstadoCivil.ShowMenu Me.ComboEstadoCivil.Left + Me.Left, Me.ComboEstadoCivil.Top + Me.Top + 230

End Sub
Private Sub ComboSexo_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 CargaLosMenus "Sexo"
 Me.MenuSexo.ShowMenu Me.ComboSexo.Left + Me.Left, Me.ComboSexo.Top + Me.Top + 230

End Sub
Private Sub ComboSigno_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 CargaLosMenus "Signo"
 Me.MenuSigno.ShowMenu Me.ComboSigno.Left + Me.Left, Me.ComboSigno.Top + Me.Top + 230

End Sub
Public Sub CargarTextos()

 ' **************************************************************
 ' Carga los Textos del Formulario
 ' **************************************************************
 ' Titulo Ventna
 ' **************************************************************
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 ' Labels
 Me.Label1.ForeColor = Variables.FontLabelColor
 Me.Label1 = MensajeRecurso(231)  ' Alias de Usuario:
 Me.Label2.ForeColor = Variables.FontLabelColor
 Me.Label2 = MensajeRecurso(127)  ' Apellido y Nombre:
 Me.Label3.ForeColor = Variables.FontLabelColor
 Me.Label3 = MensajeRecurso(233)  ' Direccion de E-Mail:
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(234)  ' Fechade Nacimiento
 Me.Label5.ForeColor = Variables.FontLabelColor
 Me.Label5 = MensajeRecurso(235)  ' Edad
 Me.Label6.ForeColor = Variables.FontLabelColor
 Me.Label6 = MensajeRecurso(236)  ' Sexo
 Me.Label7.ForeColor = Variables.FontLabelColor
 Me.Label7 = MensajeRecurso(237)  ' Ubicacion Geografica
 Me.Label8.ForeColor = Variables.FontLabelColor
 Me.Label8 = MensajeRecurso(238)  ' Intencion
 Me.Label9.ForeColor = Variables.FontLabelColor
 Me.Label9 = MensajeRecurso(239)  ' Humor
 Me.Label10.ForeColor = Variables.FontLabelColor
 Me.Label10 = MensajeRecurso(240) ' Ocupacion
 Me.Label11.ForeColor = Variables.FontLabelColor
 Me.Label11 = MensajeRecurso(241) ' Signo
 Me.Label12.ForeColor = Variables.FontLabelColor
 Me.Label12 = MensajeRecurso(242) ' Estado Civil
 Me.Label13.ForeColor = Variables.FontLabelColor
 Me.Label13 = MensajeRecurso(243) ' Telefono
 Me.Label14.ForeColor = Variables.FontLabelColor
 Me.Label14.ForeColor = Variables.FontLabelColor
 Me.Label14 = MensajeRecurso(244) ' Otra Info
 ' Botones...
 Me.LabelGrabar.ForeColor = Variables.FontBotonesColor
 Me.LabelGrabar = MensajeRecurso(245) ' Grabar
 Me.LabelOkLBL.ForeColor = Variables.FontBotonesColor
 Me.LabelOkLBL = MensajeRecurso(106) ' Cancelar
 Me.LabelRefrecar.ForeColor = Variables.FontBotonesColor
 Me.LabelRefrecar = MensajeRecurso(247) ' Refrescar...
 Me.ShapeGrabar.BackColor = Variables.ShapesBackColor
 Me.ShapeOk.BackColor = Variables.ShapesBackColor
 Me.ShapeRefrescar.BackColor = Variables.ShapesBackColor
 Me.ShapeGrabar.BorderColor = Variables.ShapesBorderColor
 Me.ShapeOk.BorderColor = Variables.ShapesBorderColor
 Me.ShapeRefrescar.BorderColor = Variables.ShapesBorderColor
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 Me.ComboSexo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.ComboSigno.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.ComboEstadoCivil.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture

End Sub
' **************************************************************
Sub MenuSigno_Click(ByVal Index As Long, Tag As String)

 Me.Signo = Tag
 
End Sub
Sub MenuEstadoCivil_Click(ByVal Index As Long, Tag As String)

 Me.EstadoCivil = Tag
 
End Sub
Sub MenuSexo_Click(ByVal Index As Long, Tag As String)

 Me.Sexo = Tag
 
End Sub
Private Sub CargaLosMenus(MenuNombre As String)

 Select Case MenuNombre
  Case "Sexo":
   ' **************************************************************
   ' Carga El Menu de Sexo
   ' **************************************************************
   Set MenuSexo = New IcoMenu
   With MenuSexo
    ' Masculino
    .SetItem 0, MensajeRecurso(248), Cliente.Imagenes.ListImages("Hombre").Picture, MensajeRecurso(248)
    ' Femenino
    .SetItem 1, MensajeRecurso(249), Cliente.Imagenes.ListImages("Mujer").Picture, MensajeRecurso(249)
   End With
 
  Case "EstadoCivil":
   ' **************************************************************
   ' Carga El Menu de Estado Civil
   ' **************************************************************
   Set MenuEstadoCivil = New IcoMenu
   With MenuEstadoCivil
    ' Casado
    .SetItem 0, MensajeRecurso(250), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(250)
    ' Divorsiado
    .SetItem 1, MensajeRecurso(251), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(251)
    ' Soltero
    .SetItem 2, MensajeRecurso(252), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(252)
    ' Viudo
    .SetItem 3, MensajeRecurso(253), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(253)
   End With
 
  Case "Signo":
   ' **************************************************************
   ' Carga El Menu De Signos
   ' **************************************************************
   Set MenuSigno = New IcoMenu
   With MenuSigno
    ' Capricornio
    .SetItem 0, MensajeRecurso(254), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(254)
    ' Acuario
    .SetItem 1, MensajeRecurso(255), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(255)
    ' Pisis
    .SetItem 2, MensajeRecurso(256), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(256)
    ' Aries
    .SetItem 3, MensajeRecurso(257), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(257)
    ' Tauro
    .SetItem 4, MensajeRecurso(258), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(258)
    ' Geminis
    .SetItem 5, MensajeRecurso(259), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(259)
    ' Cancer
    .SetItem 6, MensajeRecurso(260), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(260)
    ' Leo
    .SetItem 7, MensajeRecurso(261), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(261)
    ' Virgo
    .SetItem 8, MensajeRecurso(262), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(262)
    ' Libra
    .SetItem 9, MensajeRecurso(263), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(263)
    ' Escorpio
    .SetItem 10, MensajeRecurso(264), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(264)
    ' Sagitario
    .SetItem 11, MensajeRecurso(265), Cliente.Imagenes.ListImages("Grupo").Picture, MensajeRecurso(265)
   End With
  
 End Select
 
End Sub

Private Sub Animacion_Timer()

 If Animacion = False Then Exit Sub
 
 ' Verifica que figura debe mostrar
 IndiceAnimacion = IndiceAnimacion + 1
 If IndiceAnimacion = 5 Then IndiceAnimacion = 2
 
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(IndiceAnimacion).Picture

End Sub

Private Sub BotonGrabar_Click()
Dim PaqueteDatos As String
Dim Sexo, EstadoCivil, Signo As String
Dim Contador As Integer

 ' **************************************************************
 ' Esta Conectado?
 ' **************************************************************
 If Configuracion.Logueado <> 3 Then
  MostrarMSGBox MensajeRecurso(448), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Exit Sub
 End If

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeGrabar
 
 ' **************************************************************
 ' Verifica que no se este ejecutando la Accion. Si es asi sale...
 ' **************************************************************
 If Me.Grabando Then
  Exit Sub
 End If

 ' **************************************************************
 ' Define que se esta ejecutando la accion...
 ' **************************************************************
 Me.Grabando = True
 
 ' **************************************************************
 ' Comienza el Proceso de Grabacion de Informacion
 ' **************************************************************
 CronometroInicio = Time
 Animacion.Enabled = True
 GrabacionTimeOut.Enabled = True
  
 ' **************************************************************
 ' Completa los Especiales
 ' **************************************************************
 Sexo = Mid$(Me.Sexo, 1, 1)
 'EstadoCivil = Mid$(Me.EstadoCivil, 1, 1)
 Select Case Me.EstadoCivil
   Case MensajeRecurso(250) '"C"
    EstadoCivil = "C"
   Case MensajeRecurso(251) '"D"
    EstadoCivil = "D"
   Case MensajeRecurso(253) '"V"
    EstadoCivil = "V"
   Case MensajeRecurso(252) '"S"
    EstadoCivil = "S"
  End Select

 ' **************************************************************
 ' Arregla el Signo del Usuario...
 ' **************************************************************
 Signo = Me.Signo
 For Contador = 0 To 11
  ' Si esta en Inlges lo Pasa a Espñol...
  If UCase(Trim(MensajeRecursoReal(1000 + 254 + Contador))) = UCase(Trim(Signo)) Then
   Signo = MensajeRecursoReal(Contador + 254)
  End If
 Next
 
 ' **************************************************************
 ' Arma el Paquete con la Info del usuario
 ' **************************************************************
 PaqueteDatos = "21" & _
                CompletarCadena(Me.ApellidoYNombre, 50, "D", " ") & _
                CompletarCadena(Me.DireccionDeEmail, 50, "D", " ") & _
                CompletarCadena(Me.Edad, 2, "D", " ") & _
                CompletarCadena(CStr(Sexo), 1, "D", " ") & _
                CompletarCadena(Me.UbicacionGeografica, 20, "D", " ") & _
                CompletarCadena(Me.Intencion, 20, "D", " ") & _
                CompletarCadena(Me.Humor, 20, "D", " ") & _
                CompletarCadena(Me.Ocupacion, 20, "D", " ") & _
                CompletarCadena(Signo, 15, "D", " ") & _
                CompletarCadena(CStr(EstadoCivil), 1, "D", " ") & _
                CompletarCadena(Me.Telefono, 50, "D", " ") & _
                CompletarCadena(Me.OtraInfo, 150, "D", " ") & _
                CompletarCadena(Me.FechaDeNacimiento, 10, "D", " ")

 ' **************************************************************
 ' Envia el Paquete
 ' **************************************************************
 EnviarPaqueteTCP PaqueteDatos
 
 ' **************************************************************
 ' Define el Nuevo sexo Solicitado...
 ' **************************************************************
 SexoTemporal = UCase(CStr(Sexo))
 
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

 Unload Me
 
End Sub
Public Sub RefrescarDatos()

  ' **************************************************************
  ' Verifica que no se este ejecutando la Accion. Si es asi sale...
  ' **************************************************************
  If Me.Refrescando Then
   Exit Sub
  End If
  
  ' **************************************************************
  ' Define que se esta ejecutando la accion...
  ' **************************************************************
  Me.Refrescando = True
  
  ' **************************************************************
  ' Hace un Refresh de los Datos del Usuario
  ' **************************************************************
   Me.BlanquearCampos
   CronometroInicio = Time
   EnviarPaqueteTCP ("20" & CompletarCadena(Me.AliasUsuario, 16, "D", " "))
   Animacion.Enabled = True
   TimeOut.Enabled = True
  
End Sub
Private Sub BotonRefrescar_Click()

 ' **************************************************************
 ' Esta Conectado?
 ' **************************************************************
 If Configuracion.Logueado <> 3 Then
  MostrarMSGBox MensajeRecurso(448), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeRefrescar

 ' **************************************************************
 ' Hace un Refresh de los Datos del Usuario
 ' **************************************************************
 RefrescarDatos
  
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 ReleaseCapture
 SendMessage Me.hwnd, &HA1, 2, 0
 Exit Sub
End If

End Sub
Public Sub BlanquearCampos()

 ' **************************************************************
 ' Pone en Estado refrescando los Campos del Formulario
 ' **************************************************************
 Me.IDAliasUsuario = Me.AliasUsuario
 ' Refrescando...
 Me.ApellidoYNombre = MensajeRecurso(266)
 ' Refrescando...
 Me.DireccionDeEmail = MensajeRecurso(266)
 ' ..
 Me.FechaDeNacimiento = MensajeRecurso(267)
 ' ..
 Me.Edad = MensajeRecurso(267)
 ' Refrescando...
 Me.Sexo = MensajeRecurso(266)
 ' Refrescando...
 Me.UbicacionGeografica = MensajeRecurso(266)
 ' Refrescando...
 Me.Intencion = MensajeRecurso(266)
 ' Refrescando...
 Me.Humor = MensajeRecurso(266)
 ' Refrescando...
 Me.Ocupacion = MensajeRecurso(266)
 ' Refrescando...
 Me.Signo = MensajeRecurso(266)
 ' Refrescando...
 Me.EstadoCivil = MensajeRecurso(266)
 ' Refrescando...
 Me.Telefono = MensajeRecurso(266)
 ' Refrescando...
 Me.OtraInfo = MensajeRecurso(266)
 ' Refrescando...
 Me.EstadoUsuarioTexto = MensajeRecurso(266)
 Me.EstadoUsuarioImagen = Cliente.Imagenes.ListImages("Refrescando").Picture
 Me.EstadoNumero = "-2"
 
End Sub
Private Sub Form_Load()

 
 ' **************************************************************
 ' Define que es formulario de Datos
 ' **************************************************************
 Me.FormularioNombre = "DatosUsuario"
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
 ' **************************************************************
 ' Carga los Menus
 ' **************************************************************
 ' CargaLosMenus
 ' No los carga ya que si lo hace en el modulo que crea el Form,
 ' como toma el ultimo como valido y al hacer el On-Load
 ' se cargan los menu desplegables (Que son Formularios) da error...
 
 ' **************************************************************
 ' Pone el Titulo de la Ventna
 ' **************************************************************
 Me.TituloVentana1 = Trim(Configuracion.TituloVentanas) & "..."
 
 ' **************************************************************
 ' Completar los Campos del Formulario
 ' **************************************************************
 Me.BlanquearCampos
 
 ' **************************************************************
 ' Define el Estado de las Acciones del Formulario
 ' **************************************************************
 Me.Grabando = False
 Me.Refrescando = False
 
 ' **************************************************************
 CronometroInicio = Time
 TimeOut.Enabled = True
 Animacion.Enabled = True
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 
End Sub
Private Sub GrabacionTimeOut_Timer()
Dim Respuesta, Contador As Integer

 ' **************************************************************
 ' Define como Time Out para la Grabacion de Datos de Usuario
 ' 5 Segundos
 ' **************************************************************
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbHourglass
 If GrabacionTimeOut Then
  ' Ok se Grabaron los Datos
  If Me.GraboLosDatosUsuario = True Then
   GrabacionTimeOut.Enabled = False
   Animacion.Enabled = False
   Me.Grabando = False
   ' Pone la Imagen de Desconectado
   AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
   ' El Cambio de sus Datos Fue Exitoso...
   MostrarMSGBox MensajeRecurso(268), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   
   ' **************************************************************
   ' Define el Sexo del Usuario en el Listado de Amigos...
   ' **************************************************************
   Configuracion.Sexo = Variables.SexoTemporal
   If UCase(Trim(Configuracion.Sexo)) = "F" Then
     Cliente.ListadoDeAmigos.Nodes(2).Image = "Mujer"
     Cliente.ListadoDeAmigos.Nodes(2).SelectedImage = "Mujer"
    Else
     Cliente.ListadoDeAmigos.Nodes(2).Image = "Hombre"
     Cliente.ListadoDeAmigos.Nodes(2).SelectedImage = "Hombre"
   End If
   ' Me.Sexo = Variables.SexoTemporal
   ' Pone el Estado para reflejar elcambio de Sexo...
   PonerElEstadoDelUsuario
   
    
  ' **************************************************************
  ' Busca si existe el Formulario de Datos de Usuario Abierto, y
  ' en tal caso como se grabo la info realiza un refresh...
  ' **************************************************************
  For Contador = 1 To Forms.Count - 1
   Dim FormularioNombre, AliasUsuario, UsuarioAliasPaquete As String
   ' **************************************************************
   ' Verifica que el Formulario sea de Datos
   ' **************************************************************
   If Forms(Contador).FormularioNombre = "DatosUsuario" Then
    FormularioNombre = Trim(Forms(Contador).FormularioNombre)
    AliasUsuario = Trim(Forms(Contador).AliasUsuario)
    UsuarioAliasPaquete = Trim(Configuracion.IDAliasUsuario)
    If FormularioNombre = "DatosUsuario" And AliasUsuario = UsuarioAliasPaquete And Forms(Contador).CambioDeDatosUsuario = False Then
     Forms(Contador).RefrescarDatos
    End If
   End If
  Next
 
 End If
  
  ' Time out
  If DateDiff("s", CronometroInicio, Time) >= Configuracion.TimeOutGeneral Then
   GrabacionTimeOut.Enabled = False
   ' No Fue Posible Grabar la Información del Usuario...
   Respuesta = MostrarMSGBox(MensajeRecurso(269), vbOKOnly, "vbCritical", Configuracion.TituloVentanas)
   ' Cancela la Peticion
   Animacion.Enabled = False
   AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
   GrabacionTimeOut.Enabled = False
   Me.Grabando = False
  End If
 
 End If
 
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbDefault
 
End Sub

Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Unload Me
 
End Sub

Private Sub TimeOut_Timer()
Dim Respuesta As Integer

 ' **************************************************************
 ' Define como Time Out para la Traida de Datos de Usuario
 ' 5 Segundos
 ' **************************************************************
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbHourglass
 If TimeOut Then
  ' Ok se trajeron los Datos
  If Me.Refresco = True Then
   TimeOut.Enabled = False
   Animacion.Enabled = False
   ' Pone la Imagen de Desconectado
   AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
   ' Define que no se esta ejecutando la accion
   Me.Refrescando = False
  End If
  
  ' Time out
  If DateDiff("s", CronometroInicio, Time) >= Configuracion.TimeOutGeneral Then
   TimeOut.Enabled = False
   ' No Fue Posible Traer la Información del Usuario % ...¿Desea Intentarlo Nuevamente?
   Respuesta = MostrarMSGBox(MensajeRecurso(270) & Trim(Me.AliasUsuario) & MensajeRecurso(271), vbYesNo, "vbCritical", Configuracion.TituloVentanas)
   If Respuesta = vbYes Then
     ' Reenviar el Paquete de Peticion de Datos
     Me.RefrescarDatos
    Else
     ' Cancela la Peticion
     Animacion.Enabled = False
     AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
     TimeOut.Enabled = False
     Me.BlanquearCampos
     Me.Refrescando = False
   End If
  End If
 End If
 
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbDefault
 
End Sub
Public Sub PonerElEstadoDelUsuario()
Dim Contador As Integer

 ' **************************************************************
 ' Refrescando...
 ' **************************************************************
 If Me.EstadoNumero = "-2" And Me.Refrescando = True Then
  Me.EstadoUsuarioImagen = Cliente.Imagenes.ListImages("Refrescando").Picture
  ' Ojo que mas abajo nuevamente define el estado Refrescando...
  Exit Sub
 End If
 
 ' **************************************************************
 ' Cuando esta trabajando con los datos propios...
 ' **************************************************************
 If UCase(Trim(Me.AliasUsuario)) = UCase(Trim(Configuracion.IDAliasUsuario)) Then
  If Configuracion.Logueado = 0 Then ' Desconectado
   Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages("Desconectado").Picture
   Me.EstadoUsuarioTexto = " - " & MensajeRecurso(228) & "..."
   Exit Sub
  End If
  If Configuracion.Logueado = 1 Then ' Conectando
   Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages("Conectando").Picture
   Me.EstadoUsuarioTexto = " - " & MensajeRecurso(229) & "..."
   Exit Sub
  End If
  Me.EstadoNumero = Configuracion.EstadoDelUsuario
  Me.EstadoUsuarioTexto = Configuracion.EstadoActualTexto
  Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages(UCase(Trim(Configuracion.Sexo)) & Me.EstadoNumero).Picture
  If Me.EstadoNumero = "0" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(287)
  If Me.EstadoNumero = "1" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(180)
  If Me.EstadoNumero = "2" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(181)
  If Me.EstadoNumero = "3" Then Me.EstadoUsuarioTexto = " - " & ArreglarLenguaje(Trim(Me.EstadoUsuarioTexto))
  Exit Sub
 End If
 
 ' **************************************************************
 ' Recorre todo los Usuario del Listado de Maigos y cuando
 ' Encuentra el NumeroDeAmigoEnChat Completa los Datos
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
   If UCase(Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(Me.AliasUsuario)) Then
    Me.EstadoNumero = UCase(Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado))
    If UCase(Trim(Me.Sexo)) = UCase(Trim(MensajeRecurso(266))) Then
      ' Estado Refrescando...
      Me.EstadoUsuarioImagen = Cliente.Imagenes.ListImages("Refrescando").Picture
      Me.EstadoUsuarioTexto = MensajeRecurso(266)
     Else
      ' Define el Sexo (Imagen)...
      If UCase(Trim(Variables.GrupoAmigo(Contador).Sexo)) <> "M" And UCase(Trim(Variables.GrupoAmigo(Contador).Sexo)) <> "F" Then
        If Me.EstadoNumero = "0" Then Me.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("NoConectado").Picture
        If Me.EstadoNumero = "1" Then Me.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoVisible").Picture
        If Me.EstadoNumero = "2" Then Me.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoNoDisponible").Picture
        If Me.EstadoNumero = "3" Then Me.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoCustom").Picture
       Else
        ' Define los estados M y F (Imagen)...
        Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages(UCase(Trim(Variables.GrupoAmigo(Contador).Sexo)) & Me.EstadoNumero).Picture
      End If
      ' Define el Texto...
      If Me.EstadoNumero = "0" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(287)
      If Me.EstadoNumero = "1" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(180)
      If Me.EstadoNumero = "2" Then Me.EstadoUsuarioTexto = " - " & MensajeRecurso(181)
      If Me.EstadoNumero = "3" Then Me.EstadoUsuarioTexto = " - " & ArreglarLenguaje(Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoTexto))
    End If
    Exit For
   End If
 Next

End Sub

