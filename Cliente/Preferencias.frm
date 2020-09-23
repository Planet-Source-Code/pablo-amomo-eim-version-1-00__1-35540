VERSION 5.00
Begin VB.Form Preferencias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   ClipControls    =   0   'False
   Icon            =   "Preferencias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Preferencias.frx":000C
   ScaleHeight     =   4455
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LenguajeActual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      Top             =   1470
      Width           =   1395
   End
   Begin VB.CheckBox AvisarMensajesPendientes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2430
      TabIndex        =   34
      Top             =   3450
      Width           =   200
   End
   Begin VB.TextBox PasarAEnseguidaVuelvo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   31
      Top             =   3120
      Width           =   495
   End
   Begin VB.CheckBox CargarMinimizado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4350
      TabIndex        =   29
      Top             =   2130
      Width           =   200
   End
   Begin VB.CheckBox LogueoAutomatico 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2430
      TabIndex        =   26
      Top             =   2460
      Width           =   200
   End
   Begin VB.TextBox FontDefinidaTamano 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4590
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      Top             =   2790
      Width           =   585
   End
   Begin VB.TextBox FontDefinidaNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      Top             =   2790
      Width           =   1665
   End
   Begin VB.CheckBox ArranqueConWindows 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2430
      TabIndex        =   21
      Top             =   2130
      Width           =   200
   End
   Begin VB.TextBox TiempoDeRefresco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1140
      Width           =   465
   End
   Begin VB.TextBox DirectotioUpload 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3570
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   6
      Text            =   "Opcion No Implementada..."
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox DirectorioDownload 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4170
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   5
      Text            =   "Opcion No Implementada..."
      Top             =   780
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox SonidoActivado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2430
      TabIndex        =   4
      Top             =   1800
      Width           =   200
   End
   Begin VB.CheckBox InformaCambioDeEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2430
      TabIndex        =   3
      Top             =   1470
      Width           =   200
   End
   Begin VB.TextBox PuertoTCP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      MaxLength       =   5
      TabIndex        =   1
      Top             =   810
      Width           =   1035
   End
   Begin VB.TextBox Servidor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   3285
   End
   Begin VB.Label AsteriscoClick 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      MouseIcon       =   "Preferencias.frx":2DED
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   3450
      Width           =   2925
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   3030
      TabIndex        =   41
      Top             =   3120
      Width           =   165
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   3030
      TabIndex        =   40
      Top             =   1170
      UseMnemonic     =   0   'False
      Width           =   165
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   2670
      TabIndex        =   39
      Top             =   3450
      Width           =   165
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   2670
      TabIndex        =   38
      Top             =   1470
      UseMnemonic     =   0   'False
      Width           =   165
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   5490
      TabIndex        =   37
      Top             =   2790
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5580
      MouseIcon       =   "Preferencias.frx":2F3F
      MousePointer    =   99  'Custom
      Top             =   1470
      Width           =   240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3090
      TabIndex        =   35
      Top             =   1500
      Width           =   1125
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   33
      Top             =   3470
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3240
      TabIndex        =   32
      Top             =   3135
      Width           =   2565
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   30
      Top             =   3140
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2790
      TabIndex        =   28
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   27
      Top             =   2490
      Width           =   2385
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   5190
      MouseIcon       =   "Preferencias.frx":3091
      MousePointer    =   99  'Custom
      Top             =   2790
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4140
      MouseIcon       =   "Preferencias.frx":31E3
      MousePointer    =   99  'Custom
      Top             =   2790
      Width           =   240
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   25
      Top             =   2810
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   22
      Top             =   2160
      Width           =   2265
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   30
      TabIndex        =   19
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   14
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   15
      Top             =   1830
      Width           =   2355
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   20
      Top             =   120
      Width           =   3315
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3210
      TabIndex        =   18
      Top             =   1170
      Width           =   1995
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   17
      Top             =   1185
      Width           =   2505
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   16
      Top             =   510
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   13
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   5375
      MouseIcon       =   "Preferencias.frx":3335
      MousePointer    =   99  'Custom
      Top             =   75
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   5730
      MouseIcon       =   "Preferencias.frx":3487
      MousePointer    =   99  'Custom
      Top             =   75
      Width           =   240
   End
   Begin VB.Label BotonAplicar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2235
      MouseIcon       =   "Preferencias.frx":35D9
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3870
      Width           =   1575
   End
   Begin VB.Label BotonOk 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   4170
      MouseIcon       =   "Preferencias.frx":372B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3900
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
      Left            =   4200
      TabIndex        =   10
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Label BotonCancelar 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "Preferencias.frx":387D
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3870
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
      Left            =   255
      TabIndex        =   9
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   3880
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   4185
      Shape           =   4  'Rounded Rectangle
      Top             =   3880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   4005
      Width           =   75
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2235
      Shape           =   4  'Rounded Rectangle
      Top             =   3880
      Width           =   1575
   End
End
Attribute VB_Name = "Preferencias"
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

Public WithEvents MenuDeTamanos As IcoMenu
Attribute MenuDeTamanos.VB_VarHelpID = -1
Public WithEvents MenuDeLenguaje As IcoMenu
Attribute MenuDeLenguaje.VB_VarHelpID = -1

Dim ConfiguracionTemp As MiConfiguracion

Option Explicit
Private Function VerificarCambios() As Boolean
Dim CambioAlgo, CambioTemp As Boolean

 ' **************************************************************
 ' Cambio Algo?
 ' **************************************************************
 CambioAlgo = False
  
 ' **************************************************************
 ' Verifica si Cambio Algo...
 ' **************************************************************
 If UCase(Trim(Me.Servidor)) <> UCase(Trim(ConfiguracionTemp.Servidor)) Then CambioAlgo = True
 If Me.PuertoTCP <> ConfiguracionTemp.PortTCP Then CambioAlgo = True
 ' Sonido Activado...
 If Me.SonidoActivado = 1 Then
   CambioTemp = True
  Else
   CambioTemp = False
 End If
 If CambioTemp <> Configuracion.SonidoActivado Then CambioAlgo = True
 
 If Configuracion.Logueado = 3 Then
  If Me.TiempoDeRefresco <> ConfiguracionTemp.TiempoDeRefrescoAmigos Then CambioAlgo = True
  If Me.PasarAEnseguidaVuelvo <> ConfiguracionTemp.TiempoParaPasaraInactivo Then CambioAlgo = True
  If Configuracion.TiempoDeRefrescoAmigos <> Me.TiempoDeRefresco Then CambioAlgo = True
  ' Cambio de Estado...
  If Me.InformaCambioDeEstado = 1 Then
    CambioTemp = True
   Else
    CambioTemp = False
  End If
  If CambioTemp <> Configuracion.InformarCambiosDeEstado Then CambioAlgo = True
  ' Avisar Mensajes Pendientes...
  If Me.AvisarMensajesPendientes = 1 Then
    CambioTemp = True
   Else
    CambioTemp = False
  End If
  If CambioTemp <> Configuracion.AvisarMensajesPendientes Then CambioAlgo = True
  If UCase(Trim(Configuracion.DirectorioDownload)) <> UCase(Trim(Me.DirectorioDownload)) Then CambioAlgo = True
  If UCase(Trim(Configuracion.DirectorioUpload)) <> UCase(Trim(Me.DirectotioUpload)) Then CambioAlgo = True
  If UCase(Trim(Configuracion.FontEstandarNombre)) <> UCase(Trim(Me.FontDefinidaNombre)) Then CambioAlgo = True
  If Configuracion.FontEstandarTamano <> Me.FontDefinidaTamano Then CambioAlgo = True
 End If
 
 ' Logueo Automatico...
 If Me.LogueoAutomatico = 1 Then
   CambioTemp = True
  Else
   CambioTemp = False
 End If
 If CambioTemp <> Configuracion.LogueoAutomatico Then CambioAlgo = True
 ' Arranque Con Windows...
 If Me.ArranqueConWindows = 1 Then
   CambioTemp = True
  Else
   CambioTemp = False
 End If
 If CambioTemp <> Configuracion.ArranqueConWindows Then CambioAlgo = True
 ' Cargar Minimizado...
 If Me.CargarMinimizado = 1 Then
   CambioTemp = True
  Else
   CambioTemp = False
 End If
 If CambioTemp <> Configuracion.CargarMinimizado Then CambioAlgo = True
 If UCase(Trim(Configuracion.Lenguaje)) <> UCase(Trim(Me.LenguajeActual)) Then CambioAlgo = True

 ' **************************************************************
 ' Devuelve si se cambio Algo o no...
 ' **************************************************************
 VerificarCambios = CambioAlgo
 

End Function
Private Sub GrabaConfig()
Dim Respuesta As Long

 ' **************************************************************
 ' Verifica los Datos...
 ' **************************************************************
 ' Servidor
 If Trim(Me.Servidor) = "" Then
  ' El Nombre del Servidor de Coneccion no es Correcto...
  MostrarMSGBox MensajeRecurso(434), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Me.Servidor = ConfiguracionTemp.Servidor
  Exit Sub
 End If
 ' Puerto TCP
 If Not IsNumeric(Me.PuertoTCP) Then
  ' El Puerto TCP Ingresado no es Valido...
  MostrarMSGBox MensajeRecurso(435), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Me.PuertoTCP = ConfiguracionTemp.PortTCP
  Exit Sub
 End If
 If Me.PuertoTCP < 1 Or Me.PuertoTCP > 64000 Then
  Me.PuertoTCP = ConfiguracionTemp.PortTCP
  ' El Puerto TCP Ingresado no es Valido...
  MostrarMSGBox MensajeRecurso(435), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Solo si esta logueado...
 ' **************************************************************
 If Configuracion.Logueado = 3 Then
  ' Tiempo de Refresco
  If Not IsNumeric(Me.TiempoDeRefresco) Then
   ' El Tiempo de Refresco no es Valido...
   MostrarMSGBox MensajeRecurso(436), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Me.TiempoDeRefresco = ConfiguracionTemp.TiempoDeRefrescoAmigos
   Exit Sub
  End If
  ' Pasar a Inactivo
  If Not IsNumeric(Me.PasarAEnseguidaVuelvo) Then
   ' El Tiempo para Pasar a 'Enseguida Vuelvo' no es Valido...
   MostrarMSGBox MensajeRecurso(437), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Me.PasarAEnseguidaVuelvo = ConfiguracionTemp.TiempoParaPasaraInactivo
   Exit Sub
  End If
 End If
 ' Sonido
 If Me.SonidoActivado = 1 Then
   Configuracion.SonidoActivado = True
  Else
   Configuracion.SonidoActivado = False
 End If
 ' **************************************************************
 ' Cambia el Icono de Sonido del Formulario Cliente...
 ' **************************************************************
 DefinirEstadoSonido

 ' **************************************************************
 ' Pasa los Datos al Sistema
 ' **************************************************************
 If Configuracion.Logueado = 3 Then
  Configuracion.TiempoParaPasaraInactivo = Me.PasarAEnseguidaVuelvo
  Configuracion.TiempoDeRefrescoAmigos = Me.TiempoDeRefresco
  If Me.InformaCambioDeEstado = 1 Then
    Configuracion.InformarCambiosDeEstado = True
   Else
    Configuracion.InformarCambiosDeEstado = False
  End If
  If Me.AvisarMensajesPendientes = 1 Then
    Configuracion.AvisarMensajesPendientes = True
   Else
    Configuracion.AvisarMensajesPendientes = False
  End If
  Configuracion.DirectorioDownload = Me.DirectorioDownload
  Configuracion.DirectorioUpload = Me.DirectotioUpload
  Configuracion.FontEstandarNombre = Me.FontDefinidaNombre
  Configuracion.FontEstandarTamano = Me.FontDefinidaTamano
 End If
 
 Configuracion.Servidor = Me.Servidor
 Configuracion.PortTCP = Me.PuertoTCP
 If Me.LogueoAutomatico = 1 Then
   Configuracion.LogueoAutomatico = True
  Else
   Configuracion.LogueoAutomatico = False
 End If
 If Me.ArranqueConWindows = 1 Then
   Configuracion.ArranqueConWindows = True
  Else
   Configuracion.ArranqueConWindows = False
 End If
 If Me.CargarMinimizado = 1 Then
   Configuracion.CargarMinimizado = True
  Else
   Configuracion.CargarMinimizado = False
 End If
 Configuracion.Lenguaje = Me.LenguajeActual
 
 ' **************************************************************
 ' Verifica si los cambios se realizaron en Coneccion, entonces
 ' Lo Informa...
 ' **************************************************************
 If ConfiguracionTemp.Servidor <> Configuracion.Servidor Or ConfiguracionTemp.PortTCP <> Configuracion.PortTCP Then
   ' Ha Cambiado la Configuracion de Conección... Para que se Hagan Efectivos sus Cambios debe Reconectarse al Servidor EIM...¿Desea Hacerlo?
   Respuesta = MostrarMSGBox(MensajeRecurso(438), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then
     Exit Sub
    Else
     ' **************************************************************
     ' Desconecta y Reconecta
     ' **************************************************************
     ' Desconecta al Cliente
     Cliente.TCPSocket.Close
     ' Pasa a Estado Desconectado
     SocketTCP.CambiarEstadoDelCliente (0)
     ' Cierra el Forumlario de Preferencias
     Unload Me
     '  Abre el Loguin
    If Variables.FormularioLoguin = False Then
     Load Loguin
     Variables.FormularioLoguin = True
    End If
    Variables.BringWindowToTop (Loguin.hwnd)
    Loguin.Show vbModal
   End If
 End If
 
 
 ' **************************************************************
 ' Carga la Configuracion en una Variable Temporal
 ' **************************************************************
 ConfiguracionTemp = Configuracion
 
End Sub

Public Sub CargarTextos()

 ' **************************************************************
 ' Carga los textos
 ' **************************************************************
 Me.Label22.ForeColor = Variables.FontLabelColor
 Me.Label21.ForeColor = Variables.FontLabelColor
 'Me.Label19.ForeColor = Variables.FontLabelColor
 Me.Label20.ForeColor = Variables.FontLabelColor
 Me.Label18.ForeColor = Variables.FontLabelColor
 Me.Label8.ForeColor = Variables.FontLabelColor
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 Me.Label6 = MensajeRecurso(419)
 Me.Label6.ForeColor = Variables.FontLabelColor
 Me.Label1 = MensajeRecurso(420)
 Me.Label1.ForeColor = Variables.FontLabelColor
 Me.Label9 = MensajeRecurso(421)
 Me.Label9.ForeColor = Variables.FontLabelColor
 Me.Label10 = MensajeRecurso(422)
 Me.Label10.ForeColor = Variables.FontLabelColor
 Me.Label3 = MensajeRecurso(423)
 Me.Label3.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(424)
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label11 = MensajeRecurso(425)
 Me.Label11.ForeColor = Variables.FontLabelColor
 Me.Label14 = MensajeRecurso(426)
 Me.Label14.ForeColor = Variables.FontLabelColor
 Me.Label13 = MensajeRecurso(427)
 Me.Label13.ForeColor = Variables.FontLabelColor
 Me.Label12 = MensajeRecurso(428)
 Me.Label12.ForeColor = Variables.FontLabelColor
 Me.Label15 = MensajeRecurso(429)
 Me.Label15.ForeColor = Variables.FontLabelColor
 Me.Label17 = MensajeRecurso(430)
 Me.Label17.ForeColor = Variables.FontLabelColor
 Me.Label16 = MensajeRecurso(431)
 Me.Label16.ForeColor = Variables.FontLabelColor
 Me.Label7 = MensajeRecurso(432)
 Me.Label7.ForeColor = Variables.FontLabelColor
  
 Me.AsteriscoClick.ForeColor = Variables.FontHipervinculoColor
 Me.AsteriscoClick = MensajeRecurso(481)
 
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.CancelarLbl = MensajeRecurso(106)
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.Label2 = MensajeRecurso(275)
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Shape6.BackColor = Variables.ShapesBackColor
 Me.Shape6.BorderColor = Variables.ShapesBorderColor
 Me.Label5 = MensajeRecurso(433)
 Me.Label5.ForeColor = Variables.FontBotonesColor
 '
 Me.Image2.Picture = Cliente.Imagenes.ListImages("Minimizar").Picture
 Me.Image5.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.Image3.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.Image4.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.Image1.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 
End Sub
Private Sub AsteriscoClick_Click()

 MostrarMSGBox MensajeRecurso(482), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 
End Sub
Private Sub Form_Unload(Cancel As Integer)

 ' **************************************************************
 ' Avisa que se Cambio el Lenguaje...
 ' **************************************************************
 ' El Lenguaje Original Fue Cambiado... Para que los Cambios se Apliquen, Deberá Reiniciar la Aplicación...
 If UCase(Trim(Variables.LenguajeActual)) <> UCase(Trim(Configuracion.Lenguaje)) Then
  MostrarMSGBox MensajeRecurso(445), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 End If

End Sub
Private Sub Image1_Click()

  ' **************************************************************
  ' Muestra el Formulario de Lenguaje
  ' **************************************************************
  Audio.EjecutarSonido "003"
  MenuDeLenguaje.ShowMenu Me.Image1.Left + Me.Left, Me.Image1.Top + Me.Top + 230

End Sub

Private Sub MenuDeLenguaje_Click(ByVal Index As Long, Tag As String)
    
   LenguajeActual = Tag
   
End Sub
Private Sub MenuDeTamanos_Click(ByVal Index As Long, Tag As String)
    
  'Me.MensajeEnviar.SelFontSize = 8 + Index * 2
  Me.FontDefinidaTamano = 8 + Index * 2

End Sub
Private Sub CargarConfig()
 
 ' **************************************************************
 ' Activa o Desactiva los Controles....
 ' **************************************************************
 If Configuracion.Logueado = 3 Then
    ' Conectado...
    PasarAEnseguidaVuelvo.Enabled = True
    TiempoDeRefresco.Enabled = True
    'SonidoActivado.Enabled = True
    AvisarMensajesPendientes.Enabled = True
    InformaCambioDeEstado.Enabled = True
    FontDefinidaNombre.Enabled = True
    FontDefinidaTamano.Enabled = True
   Else
    ' Cualquier otro estado...
    PasarAEnseguidaVuelvo.Enabled = False
    TiempoDeRefresco.Enabled = False
    'SonidoActivado.Enabled = False
    AvisarMensajesPendientes.Enabled = False
    InformaCambioDeEstado.Enabled = False
    FontDefinidaNombre.Enabled = False
    FontDefinidaTamano.Enabled = False
    PasarAEnseguidaVuelvo = ""
    TiempoDeRefresco = ""
    'SonidoActivado = 0
    AvisarMensajesPendientes = 0
    InformaCambioDeEstado = 0
    FontDefinidaNombre = ""
    FontDefinidaTamano = ""
    Me.DirectorioDownload = ""
    Me.DirectotioUpload = ""
 End If
 
 ' **************************************************************
 ' Avisa que se Cambio el Lenguaje...
 ' **************************************************************
 ' El Lenguaje Original Fue Cambiado... Para que los Cambios se Apliquen, Deberá Reiniciar la Aplicación...
 If UCase(Trim(Variables.LenguajeActual)) <> UCase(Trim(Configuracion.Lenguaje)) Then
  MostrarMSGBox MensajeRecurso(445), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 End If
 
 ' **************************************************************
 ' Carga Menu de Lenguajes
 ' **************************************************************
 Set MenuDeLenguaje = New IcoMenu
  With MenuDeLenguaje
   .SetItem 0, MensajeRecurso(439), Cliente.Imagenes.ListImages("EstadoVisible").Picture, MensajeRecurso(439)
   .SetItem 1, MensajeRecurso(440), Cliente.Imagenes.ListImages("EstadoVisible").Picture, MensajeRecurso(440)
  End With
  
 ' **************************************************************
 ' Carga Menu de Tamaños
 ' **************************************************************
 Set MenuDeTamanos = New IcoMenu
  With MenuDeTamanos
   .SetItem 0, "8" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 1, "10" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 2, "12" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 3, "14" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 4, "16" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 5, "18" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 6, "20" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 7, "22" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 8, "24" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
  End With
 
 ' **************************************************************
 ' Carga la Configuracion en una Variable Temporal
 ' **************************************************************
 ConfiguracionTemp = Configuracion
 
 ' **************************************************************
 ' Carga los Datos de la Configuracion...
 ' **************************************************************
 ' **************************************************************
 ' Esta info la Muestra solo si esta logueado...
 ' **************************************************************
 If Configuracion.Logueado = 3 Then
  Me.TiempoDeRefresco = Trim(Configuracion.TiempoDeRefrescoAmigos)
  If Configuracion.InformarCambiosDeEstado Then
    Me.InformaCambioDeEstado = 1
   Else
    Me.InformaCambioDeEstado = 0
  End If
  If Configuracion.AvisarMensajesPendientes Then
    Me.AvisarMensajesPendientes = 1
   Else
    Me.AvisarMensajesPendientes = 0
  End If
  Me.PasarAEnseguidaVuelvo = Trim(Configuracion.TiempoParaPasaraInactivo)
  Me.DirectorioDownload = Trim(Configuracion.DirectorioDownload)
  Me.DirectotioUpload = Trim(Configuracion.DirectorioUpload)
  Me.FontDefinidaNombre = Configuracion.FontEstandarNombre
  Me.FontDefinidaTamano = Configuracion.FontEstandarTamano
 End If
 
 If Configuracion.SonidoActivado Then
   Me.SonidoActivado = 1
  Else
   Me.SonidoActivado = 0
 End If
 Me.Servidor = Trim(Configuracion.Servidor)
 Me.PuertoTCP = Trim(Configuracion.PortTCP)
 Me.LenguajeActual = Trim(Configuracion.Lenguaje)
 If Configuracion.ArranqueConWindows Then
   Me.ArranqueConWindows = 1
  Else
   Me.ArranqueConWindows = 0
 End If
 If Configuracion.LogueoAutomatico Then
   Me.LogueoAutomatico = 1
  Else
   Me.LogueoAutomatico = 0
 End If
 If Configuracion.CargarMinimizado Then
   Me.CargarMinimizado = 1
  Else
   Me.CargarMinimizado = 0
 End If
 
End Sub

Private Sub BotonAplicar_Click()

 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape6
 
 ' **************************************************************
 ' Si no cambio Nada... No Graba...
 ' **************************************************************
 If VerificarCambios = False Then Exit Sub
 
 ' **************************************************************
 ' Graba los Cambios
 ' **************************************************************
 GrabaConfig
 If Configuracion.Logueado = 3 Then
   Inicializar.GrabarConfiguracion
  Else
   Inicializar.GrabarConfiguracion "NO"
 End If
 
End Sub

Private Sub BotonCancelar_Click()
Dim Respuesta As Long

  Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.CancelarBt
  
  ' **************************************************************
 ' Si no cambio Nada... No Graba...
 ' **************************************************************
 If VerificarCambios = True Then
  ' **************************************************************
  ' Cancela el Cambio de Preferencias
  ' **************************************************************
  ' Al Salir por Aqui, perderá Todos los Cambios Realizados... ¿Desea Continuar?
  Respuesta = MostrarMSGBox(MensajeRecurso(441), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
  If Respuesta = vbNo Then Exit Sub ' Sale...
 End If
 
 ' Sale....
 Unload Me
 
End Sub
Public Sub CambiarLetraRemoto(Letras As String)

 ' **************************************************************
 ' Setea el Nuevo Font
 ' **************************************************************
 If Letras <> "" Then
  Me.FontDefinidaNombre = Letras
 End If
  
End Sub
Private Sub BotonOk_Click()

 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape3
 
 ' **************************************************************
 ' Si no cambio Nada... No Graba...
 ' **************************************************************
 If VerificarCambios = True Then
  ' **************************************************************
  ' Graba los Cambios
  ' **************************************************************
  GrabaConfig
  If Configuracion.Logueado = 3 Then
    Inicializar.GrabarConfiguracion
   Else
    Inicializar.GrabarConfiguracion "NO"
  End If
 End If
 
 Unload Me
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 ReleaseCapture
 SendMessage Me.hwnd, &HA1, 2, 0
 Exit Sub
End If

End Sub
Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos...
 ' **************************************************************
 Me.CargarTextos

 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(442)
 Me.Caption = TituloVentana1
 CargarConfig
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Icon = Cliente.Icon
 
End Sub
Private Sub Image2_Click()

 Audio.EjecutarSonido "003"
 Me.WindowState = vbMinimized
 
End Sub

Private Sub Image3_Click()

Dim TiposDeLetras As New EleccionTipoDeLetra

 ' **************************************************************
 ' Si no esta Logueado no da Bola
 ' **************************************************************
 If Configuracion.Logueado = 0 Then Exit Sub
 
 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Habre el Formulario de Letras
 ' **************************************************************
 Set TiposDeLetras = New EleccionTipoDeLetra
 With TiposDeLetras
  ' Le Define cual es la FontActual
  .MostrarFormulario Me.FontDefinidaNombre, Me.hwnd, False
  ' Pone la nueva Font si es "" es por que se cancelo
 End With


End Sub

Private Sub Image4_Click()

  ' **************************************************************
  ' Si no esta Logueado no da Bola
  ' **************************************************************
  If Configuracion.Logueado = 0 Then Exit Sub

  ' **************************************************************
  ' Muestra el Formulariode Tamanios de Letra
  ' **************************************************************
  Audio.EjecutarSonido "003"
  MenuDeTamanos.ShowMenu Me.Image4.Left + Me.Left, Me.Image4.Top + Me.Top + 230
  
End Sub

Private Sub Image5_Click()

 ' **************************************************************
 ' Sale
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Unload Me
 
End Sub

