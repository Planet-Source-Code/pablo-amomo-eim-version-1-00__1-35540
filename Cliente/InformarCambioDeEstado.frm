VERSION 5.00
Begin VB.Form InformarCambioDeEstado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   Icon            =   "InformarCambioDeEstado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "InformarCambioDeEstado.frx":000C
   ScaleHeight     =   1005
   ScaleMode       =   0  'User
   ScaleWidth      =   2220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer EfectoPersiana 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1230
      Top             =   990
   End
   Begin VB.Timer Timer 
      Interval        =   25000
      Left            =   1740
      Tag             =   "5000"
      Top             =   1020
   End
   Begin VB.Image ImagenEstado 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   170
      Top             =   170
      Width           =   300
   End
   Begin VB.Label UsuarioEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   690
      Width           =   2115
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1965
   End
   Begin VB.Label UsuarioNombre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   160
      Width           =   1635
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "InformarCambioDeEstado"
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
Private TamanioFinal, Saltos As Integer
Private EstadoPersiana As String
Option Explicit
Private Sub EfectoPersiana_Timer()

 If EfectoPersiana Then
  ' **************************************************************
  ' Define la Apertura
  ' **************************************************************
  If EstadoPersiana = "Abrir" Then
   Me.HeighT = Me.HeighT + Saltos
   Me.Top = Me.Top - Saltos
   ' El formulario tiene el Tamaño esperado, por lo que cancela la expancion
   If Me.HeighT >= TamanioFinal Then
    Me.EfectoPersiana.Enabled = False
   End If
  End If
  ' **************************************************************
  ' El Cierre
  ' **************************************************************
  If EstadoPersiana = "Cerrar" Then
   If Me.HeighT = 0 Then
    EfectoPersiana.Enabled = False
    ' Exit Sub
    Unload Me
   End If
   Me.HeighT = Me.HeighT - Saltos
   Me.Top = Me.Top + Saltos
   ' El formulario tiene el Tamaño esperado, por lo que cancela la expancion
   If Me.HeighT <= 0 Then
    EfectoPersiana.Enabled = False
    ' exit sub
    Unload Me
   End If
  End If
 End If
 
End Sub

Public Sub CargarTextos()

 ' ************************************************************
 ' Carga los Textos del Formulario
 ' ************************************************************
 Me.FormularioNombre = "CambioDeEstado"
 Me.Label2.ForeColor = Variables.FontCambioDeEstado2
 Me.Label2 = MensajeRecurso(286) ' Cambio su Estado a
 Me.UsuarioEstado.ForeColor = Variables.FontCambioDeEstado1
 Me.UsuarioNombre.ForeColor = Variables.FontCambioDeEstado1
 
End Sub

Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos del Formulario...
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 
End Sub

Private Sub Timer_Timer()

 If Timer Then
  Audio.EjecutarSonido "002"
  EstadoPersiana = "Cerrar"
  Me.EfectoPersiana.Enabled = True
  Timer.Enabled = False
 End If
 
End Sub
Public Sub MostrarFormulario(NombreDelUsuario As String, Estado As String, Estadotexto As String, Sexo As String)
Dim Tamanio As Integer
Dim Contado, PonerArriba, PonerAbajo As Integer
Dim Resultado As Long
Dim Area As RECT
Dim Contador, Contador2 As Integer
Dim Bandera As Boolean

 ' **************************************************************
 ' En caso de Error.... El punto es que cuando se
 ' muestra un Form VBModal, solo se puede mostrar otro
 ' vbmodal...
 ' **************************************************************
 On Error GoTo errorMostrarFormulario
 
 ' **************************************************************
 ' Carga los Datos del Formulario
 ' **************************************************************
 Me.UsuarioNombre = Trim(NombreDelUsuario)
 Select Case Trim(Estado)
  Case "0"
   ' No Conectado...
   Me.UsuarioEstado = MensajeRecurso(287)
   Me.ImagenEstado = Cliente.ImagenesAmigos.ListImages(Trim(Sexo) & Estado).Picture
  Case "1"
   ' Disponible (Normal)...
   Me.UsuarioEstado = MensajeRecurso(180)
   Me.ImagenEstado = Cliente.ImagenesAmigos.ListImages(Trim(Sexo) & Estado).Picture
  Case "2"
   ' No Disponible...
   Me.UsuarioEstado = MensajeRecurso(181)
   Me.ImagenEstado = Cliente.ImagenesAmigos.ListImages(Trim(Sexo) & Estado).Picture
  Case "3"
   Me.UsuarioEstado = Trim(Estadotexto)
   Me.ImagenEstado = Cliente.ImagenesAmigos.ListImages(Trim(Sexo) & Estado).Picture
 End Select
  
 ' **************************************************************
 ' Ojo Asegurarse que el Formulario Mide un Multiplo de 10...
 ' **************************************************************
 TamanioFinal = 1000
 Saltos = 50
 Me.EfectoPersiana.Interval = 10
 Me.Timer.Enabled = False
 Me.Timer.Interval = 5000
 Me.Timer.Enabled = True
 
 ' **************************************************************
 ' Activa la Persiana
 ' **************************************************************
 ' Define el Tamanio del WorkSpace
 Resultado = SystemParametersInfo(SPI_GETWORKAREA, 0&, Area, 0&)
 
 ' **************************************************************
 ' Lo ubica arriba del Ultimo Abierto...
 ' **************************************************************
 ' Valida todas las posiciones posibles, y verifica de abajo hacia arriba
 ' cual esta libre...
 For Contador = 1 To Int((Area.Bottom * Screen.TwipsPerPixelY) / TamanioFinal)
  Bandera = False
  For Contador2 = 0 To Forms.Count - 1
   ' Busca cada unos de los formularios a ver donde se encuentra...
   If Forms(Contador2).FormularioNombre = "CambioDeEstado" And Forms(Contador2).hwnd <> Me.hwnd Then
    'Debug.Print Forms(Contador2).Top
    'Debug.Print (Area.Bottom * Screen.TwipsPerPixelY) - (Contador * TamanioFinal) '- Screen.TwipsPerPixelY
    'Debug.Print Forms(Contador2).Top
    'Debug.Print (Area.Bottom * Screen.TwipsPerPixelY) - (Contador * TamanioFinal) + TamanioFinal
    PonerArriba = (Area.Bottom * Screen.TwipsPerPixelY) - (Contador * TamanioFinal)
    PonerAbajo = (Area.Bottom * Screen.TwipsPerPixelY) - ((Contador - 1) * TamanioFinal)
    Debug.Print Forms(Contador2).Top
    If Forms(Contador2).Top >= PonerArriba And Forms(Contador2).Top < PonerAbajo Then
    'If (Area.Bottom * Screen.TwipsPerPixelY) - (Contador * TamanioFinal) - Screen.TwipsPerPixelY And Forms(Contador2).Top <= (Area.Bottom * Screen.TwipsPerPixelY) - ((Contador - 1) * TamanioFinal) Then
     ' Aca quiere decir que hay una Aviso, por lo que sigue buscando...
     Bandera = True
     Exit For
    End If
   End If
  Next
  If Bandera = False Then Exit For
 Next
 
 If Bandera Then
   ' No hay ninguno Libre... Por lo que lo pone primero...
   Me.Top = (Area.Bottom * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
  Else
   Me.Top = (Area.Bottom * Screen.TwipsPerPixelY) - ((Contador - 1) * TamanioFinal) '- Screen.TwipsPerPixelY
 End If
 
 'Bandera = False
 'For Contador = Forms.Count - 1 To 0 Step -1
 ' If Forms(Contador).FormularioNombre = "CambioDeEstado" And Forms(Contador).hwnd <> Me.hwnd Then
 '  Bandera = True
 '  Exit For
 ' End If
 'Next
 'If Bandera Then
 '  Me.Top = Forms(Contador).Top
 ' Else
 '  Me.Top = (Area.Bottom * Screen.TwipsPerPixelY) '- (Variables.CantidadDeAvisosAbiertos * TamanioFinal)
 'End If
 
 ' **************************************************************
 ' Posiciona el Aviso...
 ' **************************************************************
 Me.Left = (Area.Right * Screen.TwipsPerPixelX) - Me.WidtH
 Me.HeighT = 0
 
 ' **************************************************************
 ' Tomar la posicion del Area de Trabajo
 ' **************************************************************
 Me.Show ' vbModal
 ' **************************************************************
 ' Una ves que lo Abrio, define que hay uno mas...
 ' **************************************************************
 'Variables.CantidadDeAvisosAbiertos = Variables.CantidadDeAvisosAbiertos + 1
 Audio.EjecutarSonido "001"
 EstadoPersiana = "Abrir"
 Me.EfectoPersiana.Enabled = True
 
errorMostrarFormulario:
 
End Sub
