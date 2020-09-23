VERSION 5.00
Begin VB.Form UsuariosBloqueados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "UsuariosBloqueados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "UsuariosBloqueados.frx":000C
   ScaleHeight     =   2700
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   3600
      ScaleHeight     =   1080
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   180
      ScaleHeight     =   1080
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   480
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   20
      Left            =   180
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   9
      Top             =   1550
      Width           =   3675
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   20
      Left            =   180
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   8
      Top             =   480
      Width           =   3675
   End
   Begin VB.TextBox Cantidades 
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
      TabIndex        =   5
      Top             =   1676
      Width           =   3975
   End
   Begin VB.ListBox UsuariosBloqueados 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3675
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   4065
   End
   Begin VB.Image ScrollArriba 
      Height          =   240
      Left            =   3980
      MouseIcon       =   "UsuariosBloqueados.frx":2283
      MousePointer    =   99  'Custom
      Top             =   460
      Width           =   240
   End
   Begin VB.Image ScrollAbajo 
      Height          =   240
      Left            =   3980
      MouseIcon       =   "UsuariosBloqueados.frx":23D5
      MousePointer    =   99  'Custom
      Top             =   1340
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "UsuariosBloqueados.frx":2527
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   3315
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Label BotonRefrescar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "UsuariosBloqueados.frx":2679
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   270
      MouseIcon       =   "UsuariosBloqueados.frx":27CB
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2535
      TabIndex        =   4
      Top             =   2270
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
      Left            =   255
      TabIndex        =   1
      Top             =   2270
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2535
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   1575
   End
End
Attribute VB_Name = "UsuariosBloqueados"
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

Public WithEvents MenuBLoqueo As IcoMenu
Attribute MenuBLoqueo.VB_VarHelpID = -1
Option Explicit

Private Sub BotonRefrescar_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape6
 
 CargarBloqueados
 
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
  '  - Amigos Bloqueados...
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(292)
 ' Botones...
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.Label4.ForeColor = Variables.FontBotonesColor
 Me.Label4 = MensajeRecurso(247)    ' Refrescar...
 Me.Shape6.BackColor = Variables.ShapesBackColor
 Me.Shape6.BorderColor = Variables.ShapesBorderColor
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(128)    ' Cerrar...
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.ScrollArriba.Picture = Cliente.ImagenesFlecha.ListImages("ArribaAzul").Picture
 Me.ScrollAbajo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoAzul").Picture
 ' Tool-Tip Texts...
 Me.UsuariosBloqueados.ToolTipText = MensajeRecurso(446) ' Listado de Amigos Blockeados...
 
End Sub
Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos
 ' **************************************************************
 Me.CargarTextos

 '  - Amigos Bloqueados...
 Me.Caption = Trim(Configuracion.TituloVentanas) & MensajeRecurso(292)
 Inicializar.CargarBloqueUsuarios
 CargarBloqueados
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Icon = Cliente.Icon
 
End Sub
Private Sub CargarBloqueados()
Dim Contador As Integer
Dim Texto As String
Dim BloqueadosSi, BloqueadosNo As Integer

 ' **************************************************************
 ' Carga los Amigos Bloqueados
 ' **************************************************************
 ' Limpia el Listado
 Me.UsuariosBloqueados.Clear
 ' Define que no Hay Amigos que se Puedan Bloquear - Desbloquear
 'If Variables.UsuarioBloqueadoCantidad = 0 Then
 ' Me.UsuariosBloqueados.AddItem ("No Hay Amigos Que Puedan Bloquearse / Desbloquearse...No Posee Usuarios Bloqueados...")
 ' Exit Sub
 'End If
 ' Carga todos los Usuarios...
 BloqueadosSi = 0
 BloqueadosNo = 0
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If Trim(GrupoAmigo(Contador).IDNombreDelAmigo) <> "" Then
   Texto = Trim(GrupoAmigo(Contador).IDNombreDelAmigo)
   If UsuarioEstaBloqueado(Texto) Then
     ' [Bloqueado]...
     Texto = Texto & MensajeRecurso(293)
     BloqueadosSi = BloqueadosSi + 1
    Else
     ' [No Bloqueado]...
     Texto = Texto & MensajeRecurso(294)
     BloqueadosNo = BloqueadosNo + 1
   End If
   Me.UsuariosBloqueados.AddItem (Texto)
  End If
 Next
 
 ' [ % ] Bloqueados, [ % ] No Bloqueados...
 Cantidades = MensajeRecurso(297) & BloqueadosSi & MensajeRecurso(295) & BloqueadosNo & MensajeRecurso(296)
 
End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Unload Me
 
End Sub

Private Sub Label3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape3

 Unload Me

End Sub
Function UsuarioEstaBloqueado(Nombre As String) As Boolean
Dim Contador As Integer

  ' **************************************************************
  ' Verifica si el Usuario esta Bloqueado
  ' **************************************************************
  If Variables.UsuarioBloqueadoCantidad = 0 Then
   UsuarioEstaBloqueado = False
   Exit Function
  End If
  
  ' **************************************************************
  ' Verifica si el Usuario esta Bloqueado
  ' **************************************************************
  For Contador = 1 To Variables.UsuarioBloqueadoCantidad
   If UCase(Trim(Nombre)) = UCase(Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)) Then
    ' Usuario Bloqueado
    UsuarioEstaBloqueado = True
    Exit Function
   End If
  Next
 
  ' **************************************************************
  ' Define que no esta Bloqueado
  ' **************************************************************
  UsuarioEstaBloqueado = False
  
End Function
Private Sub MenuBLoqueo_Click(ByVal Index As Long, Tag As String)
Dim Contador As Integer
Dim Respuesta As Long

 Select Case Index
  Case 0
   ' **************************************************************
   ' Procesa un Bloqueo de Usuario
   ' **************************************************************
   Respuesta = UsuarioEstaBloqueado(Trim(Tag))
   If Respuesta Then
    ' El Amigo [ % ] ya se encuentra Bloqueado...
    Respuesta = MostrarMSGBox(MensajeRecurso(134) & Trim(Tag) & MensajeRecurso(299), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
    Exit Sub
   End If
   ' ¿Está Seguro que Desea Bloquear al Amigo [ % ]?...
   Respuesta = MostrarMSGBox(MensajeRecurso(300) & Trim(Tag) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then Exit Sub
   Varios.ProcesarUsuariosBloqueados "Agregar", Trim(Tag)
  Case 2
   ' **************************************************************
   ' Procesa un Desbloqueo
   ' **************************************************************
   Respuesta = UsuarioEstaBloqueado(Trim(Tag))
   If Not Respuesta Then
    ' El Amigo [ % ] no se encuentra Bloqueado...
    Respuesta = MostrarMSGBox(MensajeRecurso(134) & Trim(Tag) & MensajeRecurso(302), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
    Exit Sub
   End If
   ' ¿Está Seguro que Desea Desbloquear al Amigo [ % ]?...
   Respuesta = MostrarMSGBox(MensajeRecurso(303) & Trim(Tag) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then Exit Sub
   Varios.ProcesarUsuariosBloqueados "Sacar", Trim(Tag)
 End Select
 
 ' **************************************************************
 ' Actualiza todos los Formularios
 ' **************************************************************
 For Contador = 1 To Forms.Count - 1
  If Forms(Contador).FormularioNombre = "Mensajes" Then
   Forms(Contador).CargarAmigosMultiChat
  End If
 Next
 
 ' **************************************************************
 ' Actualiza el Listado...
 ' **************************************************************
 CargarBloqueados
 
End Sub

Private Sub ScrollAbajo_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 Scroller.ScrollListBox Me.UsuariosBloqueados, "Abajo"

End Sub

Private Sub ScrollArriba_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 Scroller.ScrollListBox Me.UsuariosBloqueados, "Arriba"
 
End Sub

Private Sub UsuariosBloqueados_Click()
Dim Nombre As String
Dim Posicion As POINTAPI
Dim Posicion1 As Integer

 ' **************************************************************
 ' Toma la Posicion del Mouse
 ' **************************************************************
 GetCursorPos Posicion

 ' **************************************************************
 ' Define el Nombre Elegigo
 ' **************************************************************
 Posicion1 = InStr(1, Trim(Me.UsuariosBloqueados.List(Me.UsuariosBloqueados.ListIndex)), "[")
 If Posicion1 <> 0 Then
   Nombre = Trim(Mid$(Trim(Me.UsuariosBloqueados.List(Me.UsuariosBloqueados.ListIndex)), 1, Posicion1 - 1))
  Else
   Nombre = Trim(Me.UsuariosBloqueados.List(Me.UsuariosBloqueados.ListIndex))
 End If
 
 ' **************************************************************
 ' Muestra el Menu
 ' **************************************************************
 CargarMenu Nombre
 Me.MenuBLoqueo.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY


End Sub
Sub CargarMenu(Usuario As String)
 
 ' **************************************************************
 ' Carga el Menu con los Estados de los Usuarios
 ' **************************************************************
  Set Me.MenuBLoqueo = New IcoMenu
  With Me.MenuBLoqueo
   ' Bloquear Amigo...
   .SetItem 0, MensajeRecurso(304), Cliente.Imagenes.ListImages("UsuarioNoExiste").Picture, Trim(Usuario)
   .SetItem 1, ""
   ' Desbloquear Amigo...
   .SetItem 2, MensajeRecurso(305), Cliente.Imagenes.ListImages("EstadoVisible").Picture, Trim(Usuario)
  End With
  ' **************************************************************

End Sub

