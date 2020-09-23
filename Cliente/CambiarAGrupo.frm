VERSION 5.00
Begin VB.Form CambiarAGrupo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "CambiarAGrupo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CambiarAGrupo.frx":000C
   ScaleHeight     =   2310
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MoverAGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   2625
   End
   Begin VB.TextBox Amigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   2835
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4005
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   3405
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3900
      MouseIcon       =   "CambiarAGrupo.frx":231F
      MousePointer    =   99  'Custom
      Top             =   1170
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "CambiarAGrupo.frx":2471
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "CambiarAGrupo.frx":25C3
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1740
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1875
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "CambiarAGrupo.frx":2715
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1740
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
      TabIndex        =   3
      Top             =   1875
      Width           =   1545
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   1740
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
      TabIndex        =   0
      Top             =   630
      Width           =   1185
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "CambiarAGrupo"
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
' Define los Menus...
' **************************************************************
Public WithEvents MenuDeGrupos As IcoMenu
Attribute MenuDeGrupos.VB_VarHelpID = -1

Option Explicit
Public CambiarUsuario As String
Public CambiarUsuarioGrupo As String
Public Sub CargarTextos()

 ' **************************************************************
 ' Define los Textos...
 ' **************************************************************
 ' Titulo Ventana...
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(149)
 ' Lables...
 Me.UsuarioLbl.ForeColor = Variables.FontLabelColor
 Me.UsuarioLbl = MensajeRecurso(144) ' Amigo:
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(145) ' Mover a:
 ' Botones...
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.CancelarLbl = MensajeRecurso(106) ' Cancelar
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(147) ' Mover a Grupo...
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 ' **************************************************************
 ' Imagenes..
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.Image2.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 
End Sub
Private Sub MenuDeGrupos_Click(ByVal Index As Long, Tag As String)

 ' **************************************************************
 ' Define el Grupo Seleccionado en el Menu Desplegable...
 ' NOTA: En el TAG del Grupo se carga el Nombre del Mismo...
 ' **************************************************************
 MoverAGrupo = Tag
 
End Sub
Public Sub MostrarFormulario(Grupo As String, Usuario As String)

 ' **************************************************************
 ' Carga las Variables necesarias para trabajar
 ' **************************************************************
 If Grupo = MensajeRecurso(148) Then Grupo = MensajeRecurso(132) ' Fuera de Grupo...
 CambiarUsuario = Usuario
 CambiarUsuarioGrupo = Grupo
 Amigo = CambiarUsuario
 MoverAGrupo = Me.CambiarUsuarioGrupo
 Me.Show ' <- Muestra el Formulario...
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Permite hacer el Move del Formulario
 ' **************************************************************
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0
  Exit Sub
 End If

End Sub
Private Sub Form_Load()
Dim Contador As Integer
 
 ' **************************************************************
 ' Carga los Textos
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
 ' **************************************************************
 ' Menu De Cambio de Grupo
 ' **************************************************************
  Set Me.MenuDeGrupos = New IcoMenu
  With Me.MenuDeGrupos
   ' Fuera de Grupo...
   .SetItem 0, MensajeRecurso(132), Cliente.ImagenesMenus.ListImages("MoverAGrupo").Picture, MensajeRecurso(132)
   ' **************************************************************
   ' Carga los Grupos
   ' **************************************************************
   Dim Cantidad As Integer
   Cantidad = 0
   For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
    If Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1) = "G" Then
     Cantidad = Cantidad + 1
     .SetItem 0 + Cantidad, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2), Cliente.ImagenesMenus.ListImages("MoverAGrupo").Picture, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2)
    End If
   Next
  End With
  
End Sub
Private Sub Image2_Click() ' Desplegable del Grupo...

 ' **************************************************************
 ' Ejecuta el Sonido de Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Me.MenuDeGrupos.ShowMenu Me.Image2.Left + Me.Left, Me.Image2.Top + Me.Top + 230

End Sub
Private Sub Image3_Click() ' Cerrar...

 ' **************************************************************
 ' Ejecuta el Sonido de Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Descargar el Formulario...
 ' **************************************************************
 Unload CambiarAGrupo
 
End Sub
Private Sub Label1_Click() ' Cancelar...

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.CancelarBt

 ' **************************************************************
 ' Descargar el Formulario...
 ' **************************************************************
 Unload CambiarAGrupo
 
End Sub
Private Sub Label3_Click()
Dim Respuesta As Integer
 
 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape3
 
 Respuesta = Varios.CambiarUsuarioDeGrupo(CambiarUsuario, MoverAGrupo, True)
 If Respuesta = 1 Then Unload Me ' Se realizo el Cambio OK ! - Se descarga el Formulario...
 
End Sub
