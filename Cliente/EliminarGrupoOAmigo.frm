VERSION 5.00
Begin VB.Form EliminarGrupoOAmigo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "EliminarGrupoOAmigo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "EliminarGrupoOAmigo.frx":000C
   ScaleHeight     =   1650
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox GrupoOUsuario 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   570
      Width           =   2265
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3930
      MouseIcon       =   "EliminarGrupoOAmigo.frx":2160
      MousePointer    =   99  'Custom
      Top             =   540
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "EliminarGrupoOAmigo.frx":22B2
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   30
      Width           =   4035
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
      Left            =   180
      TabIndex        =   4
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "EliminarGrupoOAmigo.frx":2404
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
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
      TabIndex        =   3
      Top             =   1215
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "EliminarGrupoOAmigo.frx":2556
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1080
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
      Top             =   1215
      Width           =   1545
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   1075
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1075
      Width           =   1575
   End
End
Attribute VB_Name = "EliminarGrupoOAmigo"
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

Option Explicit
' Public CambiarUsuario As String
Public TipoAmigoOGrupo, TipoTemp As String
Public WithEvents ListadoCombo As IcoMenu
Attribute ListadoCombo.VB_VarHelpID = -1
Public CantidadActual As Integer

Public Sub CargarTextos()

 ' **************************************************************
 ' Define los Textos del Formulario
 ' **************************************************************
 ' Ventana...
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 ' Botones y Labels...
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(282) & ":"         ' Borrar:
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.CancelarLbl = MensajeRecurso(106)    ' Cancelar
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(282) & "..." ' Borrar...
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.Image2.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture

End Sub
Private Sub ListadoCombo_Click(ByVal Index As Long, Tag As String)

 GrupoOUsuario = Tag
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 ReleaseCapture
 SendMessage Me.hwnd, &HA1, 2, 0
 Exit Sub
End If

End Sub
Public Sub MostrarFormulario(Tipo As String)
Dim Contador As Integer
Dim Cadena As String

 ' **************************************************************
 ' Guarda la Variable
 ' **************************************************************
 TipoAmigoOGrupo = Tipo
 '1344
 TipoTemp = Tipo
 If UCase(Trim(TipoTemp)) = "AMIGO" Then
  TipoTemp = Mid$(Trim(MensajeRecurso(344)), 1, Len(Trim(MensajeRecurso(344))) - 4)
 End If
 If UCase(Trim(TipoTemp)) = "GRUPO" Then
  TipoTemp = Trim(MensajeRecurso(460))
 End If
  
 ' **************************************************************
 ' Carga el Titulo de la Ventana
 ' **************************************************************
 '  - Borrar % ...
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(280) & " " & Trim(TipoTemp) & MensajeRecurso(281)
   
 ' **************************************************************
 ' Cambiar Titulo del Control
 ' **************************************************************
 ' Borrar % :
 Label4.Caption = MensajeRecurso(282) & " " & Trim(TipoTemp) & MensajeRecurso(283)
 ' Label del Boton...
 ' Borrar % ...
 Label2.Caption = MensajeRecurso(282) & " " & Trim(TipoTemp) & MensajeRecurso(281)
 
 ' **************************************************************
 ' Carga en el Control todos los Grupos del Listado de Amigos
 ' **************************************************************
 ' Primero lo Borra
 'Me.GrupoOUsuario.Clear
 ' Carga el Control
 ' Carga el Menu Correspondiente...
 Set ListadoCombo = New IcoMenu
  With ListadoCombo
   Dim Cantidad As Integer
   Cantidad = 0
   For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
    Select Case Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1)
     Case "G"
      If Tipo = "Grupo" Then
       ' Guarda el Primero
       If Cantidad = 0 Then
        Cadena = (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2))
       End If
       .SetItem Cantidad, (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2)), Cliente.ImagenesMenus.ListImages("GrupoEliminar").Picture, (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2))
       Cantidad = Cantidad + 1
      End If
     Case "U"
      If Tipo = "Amigo" Then
       If UCase(Trim(Cliente.ListadoDeAmigos.Nodes(Contador).key)) <> UCase("Usuario") Then
        ' Guarda el Primero
        If Cantidad = 0 Then
         Cadena = (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 4))
        End If
        .SetItem Cantidad, (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 4)), Cliente.ImagenesMenus.ListImages("AmigoEliminar").Picture, (Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 4))
        Cantidad = Cantidad + 1
       End If
      End If
    End Select
   Next
  End With
 
 
 ' **************************************************************
 ' Se posiciona en el Primer Item del Control
 ' **************************************************************
 If Cantidad > 0 Then
   Me.GrupoOUsuario = Cadena
  Else
   ' No Queda %  a Eliminar...
   Me.GrupoOUsuario = MensajeRecurso(284) & " " & Trim(TipoTemp) & MensajeRecurso(285)
 End If
 
 ' **************************************************************
 ' Define cuantos Item's Hay
 ' **************************************************************
 Me.CantidadActual = Cantidad

End Sub
Private Sub Form_Load()

 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
End Sub
Private Sub Image2_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Si no hay nada que mostrar... Sale...
 ' **************************************************************
 If Me.CantidadActual <= 0 Then Exit Sub
 
 Me.ListadoCombo.ShowMenu Me.Image2.Left + Me.Left, Me.Image2.Top + Me.Top + 230

End Sub

Private Sub Image3_Click()

' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' Cancela la Operacion
 Unload Me
 
End Sub

Private Sub Label1_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape3
 
 ' **************************************************************
 ' Cancela la Operacion
 ' **************************************************************
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

 ' **************************************************************
 ' Verifica que se haya seleccionado un Grupo o Amigo
 ' **************************************************************
 ' No Queda  %  a Eliminar...
 If Me.GrupoOUsuario = MensajeRecurso(284) & Me.TipoAmigoOGrupo & MensajeRecurso(285) Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Realiza la Operacion de Borrado
 ' **************************************************************
 Select Case Me.TipoAmigoOGrupo
  Case "Grupo"
   Varios.EliminarGrupo (Me.GrupoOUsuario)
  Case "Amigo"
   Varios.EliminarAmigo (Me.GrupoOUsuario)
 End Select
 
 ' **************************************************************
 ' Recarga el COMBO
 ' **************************************************************
 Me.MostrarFormulario Me.TipoAmigoOGrupo
 
End Sub
