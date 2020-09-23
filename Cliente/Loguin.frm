VERSION 5.00
Begin VB.Form Loguin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "Loguin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loguin.frx":000C
   ScaleHeight     =   2400
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RecordarPasswordCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   3495
      TabIndex        =   13
      Top             =   1195
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox EstadoVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      ForeColor       =   &H00000000&
      Height          =   200
      Left            =   1185
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   1195
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.Timer Animacion 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2520
      Top             =   390
   End
   Begin VB.TextBox Password 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1280
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   885
      Width           =   1620
   End
   Begin VB.TextBox Usuario 
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
      Height          =   225
      Left            =   1280
      MaxLength       =   16
      TabIndex        =   1
      Top             =   540
      Width           =   2865
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1530
      TabIndex        =   14
      Top             =   1210
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "Loguin.frx":2453
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   3885
   End
   Begin VB.Label RecordarPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   660
      MouseIcon       =   "Loguin.frx":25A5
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1430
      Width           =   3075
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   1210
      Width           =   945
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   3930
      Top             =   930
      Width           =   270
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
      TabIndex        =   8
      Top             =   120
      Width           =   3315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "Loguin.frx":26F7
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1845
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "Loguin.frx":2849
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1845
      Width           =   1575
   End
   Begin VB.Label CancelarLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   1950
      Width           =   1545
   End
   Begin VB.Shape CancelarBt 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Label PasswordLbl 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   225
      TabIndex        =   2
      Top             =   920
      Width           =   945
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   560
      Width           =   945
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1830
      Width           =   1575
   End
End
Attribute VB_Name = "Loguin"
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
' **************************************************************
' Esta variable guarda la fecha/hora para calcular el tiempo de logueo
' **************************************************************
Private TiempoLogueoInicial As Date
Private IndiceAnimacion As Integer
Public EnvioPassword As Integer
Public EnvioPasswordDireccionMail As String
Public ActivoAutologueo As Boolean

Sub BotonCancelar()
 
 ' Ojo esto tiene que esta igual en BotonCancelar...
 ' **************************************************************
 ' Verifica si el Usuario esta logueado, si no es asi, cierra
 ' el sistema
 ' **************************************************************
  ' **************************************************************
 ' Define el Estado del Formulario
 ' **************************************************************
 Variables.FormularioLoguin = False

 If Configuracion.Logueado <> 3 Then
  Cliente.TCPSocket.Close
  SocketTCP.CambiarEstadoDelCliente (0)
 End If
 
 ' **************************************************************
 ' Muestra el Formulario de Cliente
 ' **************************************************************
 Unload Me
 Cliente.Show
   
End Sub
Private Sub Cancelar_Click()
 
 BotonCancelar
 
End Sub
Private Sub Conectar_Click()

 BotonConectar
 
End Sub
Private Sub BotonConectar()
Dim SegundosTranscurridos As Integer
Dim PaqueteEnviar, PWDTemp As String

 
 ' **************************************************************
 ' Dispara el Timer de La Animacion (Comunicacion)
 ' **************************************************************
 Animacion.Enabled = True
 
 ' **************************************************************
 ' Traba los Controles mientras negocia el Logueo
 ' **************************************************************
 Usuario.Enabled = False
 Password.Enabled = False
 EstadoVisible.Enabled = False
 Label1.Enabled = False
 Label3.Enabled = False
 RecordarPassword.Enabled = False
 
 ' **************************************************************
 ' Comienza el Proceso de Logueo
 ' **************************************************************
 ' Cambia el Puntero del Mouse (Reloj de Arena)
 Me.MousePointer = vbHourglass
 ' Pone en Estado Conectando
 SocketTCP.CambiarEstadoDelCliente (1)
 ' Primero Cierra el Socket
 Cliente.TCPSocket.Close
 ' Carga la Info de la Coneccion contra el Servidor...
 Cliente.TCPSocket.RemoteHost = Trim(Configuracion.Servidor)
 Cliente.TCPSocket.RemotePort = Trim(Configuracion.PortTCP)
 ' Comienza la Coneccion contra el Servidor
 Cliente.TCPSocket.Connect
 ' Pone el Estado de Logueo en 0 (No Logueado)
 Configuracion.Logueado = 0
 
 ' **************************************************************
 ' Espera 3 Segundos para que se conecte el Socket
 ' **************************************************************
 TiempoLogueoInicial = Time
 Do Until Cliente.TCPSocket.State = sckConnected
  DoEvents
  SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 ' Verifica que el Port se halla conectado... Sino es asi sale con errror...
 If Cliente.TCPSocket.State <> sckConnected Then
  ' No fue Posible Conectarse Contra el Servidor...
  MostrarMSGBox MensajeRecurso(108), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  ' Detiene el Loguin Automatico
  If Variables.LoguinAutomatico Then
   Variables.LoguinAutomatico = False
  End If
  ' Detiene la Animacion...
  Animacion.Enabled = False
  AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
  ' **
  SocketTCP.CambiarEstadoDelCliente (0)
  Me.MousePointer = vbDefault
  ' **************************************************************
  ' DesTraba los Controles... (Antes Trabados por el Logueo)
  ' **************************************************************
  Usuario.Enabled = True
  Password.Enabled = True
  EstadoVisible.Enabled = True
  Label1.Enabled = True
  Label3.Enabled = True
  RecordarPassword.Enabled = True
  ' **************************************************************
  Exit Sub ' No Continua el logueo debido a algun error...!!!
 End If
 ' **************************************************************
 
 ' **************************************************************
 ' Comienza la Negociacion de Logueo
 ' **************************************************************
 ' Encripar la Password (OPCION NO IMPLEMENTADA...)
 PWDTemp = Encriptar(Password, Password)
 ' Arma el Paquete de Loguin
 PaqueteEnviar = "00" & _
                 CompletarCadena(Usuario, 16, "I", " ") & _
                 CompletarCadena(PWDTemp, 12, "I", " ")
 ' Define que el usuairo no esta disponible
 Dim Bandera As Boolean
 Bandera = False
 If EstadoVisible = 0 Then
   Bandera = True
   PaqueteEnviar = PaqueteEnviar & "2"
   Configuracion.EstadoDelUsuario = 2
   Configuracion.EstadoActualTexto = ""
 End If
 
 ' **************************************************************
 ' Verifica si el ultimo estado es igual al actual
 ' NOTA:    Solo ejecuta esta parte si el usuario no opto por
 '          entrar como "No Disponible..."
 ' **************************************************************
 If EstadoVisible <> 0 Then
  Dim Respuesta As Long
  Inicializar.UltimoEstado "Leer", CStr(Usuario) ' Levanta el Estado de la Registry
  ' Verifica si el Numero de estado es distindo
  If CInt(Configuracion.EstadoDelUsuario) <> CInt(Variables.UltimoEstadoDelUsuario.EstadoNumero) Then
    ' Si esta como Autologueo, no pregunta, y cambia al estado anterior automaticamente
    If Variables.Configuracion.LogueoAutomatico = False Then
       ' Pregunta si quiere cambiar al estado anterior
       Select Case Variables.UltimoEstadoDelUsuario.EstadoNumero
        Case 1: ' Disponible
         ' En su ultima sesion su Estado fue 'Disponible'... ¿Desea comenzar su sesion en este estado?...
         Respuesta = MostrarMSGBox(MensajeRecurso(109), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
        Case 2: ' No Disponible
         ' En su ultima sesion su Estado fue 'No Disponible'... ¿Desea comenzar su sesion en este estado?...
         Respuesta = MostrarMSGBox(MensajeRecurso(110), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
        Case 3: ' Custom
         ' En su ultima sesion su Estado fue ' % '... ¿Desea comenzar su sesion en este estado?...
         Respuesta = MostrarMSGBox(MensajeRecurso(111) & Varios.ArreglarLenguaje(Trim(Variables.UltimoEstadoDelUsuario.Estadotexto)) & MensajeRecurso(112), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
       End Select
     Else
       ' Aca define a Respuesta como Yes para que cambie el estado
       Respuesta = vbYes
    End If
    ' Pone el Estado Anterior
    If Respuesta = vbYes Then
     Bandera = True
     PaqueteEnviar = PaqueteEnviar & CStr(Variables.UltimoEstadoDelUsuario.EstadoNumero)
     If Variables.UltimoEstadoDelUsuario.EstadoNumero = 3 Then
      PaqueteEnviar = PaqueteEnviar & CompletarCadena(Varios.ArreglarLenguaje(Trim(Variables.UltimoEstadoDelUsuario.Estadotexto)), 20, "D", " ")
     End If
     Configuracion.EstadoDelUsuario = Variables.UltimoEstadoDelUsuario.EstadoNumero
     Configuracion.EstadoActualTexto = Varios.ArreglarLenguaje(Trim(Variables.UltimoEstadoDelUsuario.Estadotexto))
    End If
   Else
    ' Esta en un Estado Custom Distinto
    If (CInt(Configuracion.EstadoDelUsuario) = 3) And Trim(Variables.UltimoEstadoDelUsuario.Estadotexto) <> Trim(Configuracion.EstadoActualTexto) Then
     ' En su ultima sesion su Estado fue ' % '... ¿Desea comenzar su sesion en este estado?...
     Respuesta = MostrarMSGBox(MensajeRecurso(111) & Varios.ArreglarLenguaje(Trim(Variables.UltimoEstadoDelUsuario.Estadotexto)) & MensajeRecurso(112), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
     If Respuesta = vbYes Then
      Bandera = True
      PaqueteEnviar = PaqueteEnviar & CStr(Variables.UltimoEstadoDelUsuario.EstadoNumero)
      If Variables.UltimoEstadoDelUsuario.EstadoNumero = 3 Then
       PaqueteEnviar = PaqueteEnviar & CompletarCadena(Varios.ArreglarLenguaje(Trim(Variables.UltimoEstadoDelUsuario.Estadotexto)), 20, "D", " ")
       Configuracion.EstadoDelUsuario = Variables.UltimoEstadoDelUsuario.EstadoNumero
       Configuracion.EstadoActualTexto = Varios.ArreglarLenguaje((Variables.UltimoEstadoDelUsuario.Estadotexto))
      End If
     End If
    End If
  End If
 End If
 ' **************************************************************
  
 ' **************************************************************
 ' Si no paso por la Instancia Anterior lo arregla
 ' **************************************************************
 If Bandera = False Then
  PaqueteEnviar = PaqueteEnviar & "1"
  Configuracion.EstadoDelUsuario = 1
  Configuracion.EstadoActualTexto = ""
 End If
  
 ' **************************************************************
 ' Enviar Paquete de Logueo (Inlcuyendo el Estado en el cual
 ' quiere aparecer, etc., etc.
 ' **************************************************************
 EnviarPaqueteTCP (PaqueteEnviar)
  
 ' **************************************************************
 ' Deja un Bucle Para Valida el Logueo o Informar que o fue exitoso
 ' tomando tiempo de TimeOut desde Configuracion.TimeOutLogueo
 ' **************************************************************
 TiempoLogueoInicial = Time
 Do
  DoEvents
  ' Si el Logueo Fue exitoso sale del Bucle
  If Configuracion.Logueado <> 0 Then Exit Do
  
  ' Verifica que no se pase el Tiempo de TimeOut
  SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutLogueo Then Exit Do
 Loop
 
 ' **************************************************************
 ' DesTraba los Controles... (Antes Trabados por el Logueo)
 ' **************************************************************
 Usuario.Enabled = True
 Password.Enabled = True
 EstadoVisible.Enabled = True
 Label1.Enabled = True
 Label3.Enabled = True
 RecordarPassword.Enabled = True
 ' **************************************************************
 
 ' **************************************************************
 ' Avisar Si Hubo algun Problema en el Logueo
 ' **************************************************************
 Dim RespuestaLogueo As Integer
 Select Case Configuracion.Logueado
  Case 0:
   ' No Fue Posible Loguearse al Servidor...
   MostrarMSGBox MensajeRecurso(113), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   ' Si no funca el Logueo para el Logueo Automatico... Sino entra en un
   ' loop esperando el OK
   If Variables.LoguinAutomatico Then
    Variables.LoguinAutomatico = False
   End If
  Case 1:
   ' El Usuario Ingresado No Existe...
   MostrarMSGBox MensajeRecurso(114), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   ' Si no funca el Logueo para el Logueo Automatico... Sino entra en un
   ' loop esperando el OK
   If Variables.LoguinAutomatico Then
    Variables.LoguinAutomatico = False
   End If
  Case 2:
   ' La Password Ingresada No es Correcta...
   MostrarMSGBox MensajeRecurso(115), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Password.SetFocus
   ' Si no funca el Logueo para el Logueo Automatico... Sino entra en un
   ' loop esperando el OK
   If Variables.LoguinAutomatico Then
    Variables.LoguinAutomatico = False
   End If
  Case 4:
   ' Su Usuario Se Encuentra Lockeado...
   MostrarMSGBox MensajeRecurso(116), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   ' Si no funca el Logueo para el Logueo Automatico... Sino entra en un
   ' loop esperando el OK
   If Variables.LoguinAutomatico Then
    Variables.LoguinAutomatico = False
   End If
 End Select
 
 ' Detiene la Animacion...
 Animacion.Enabled = False
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 ' **
 
 ' Si no se pudo loguear se cierra el Socket...
 If Configuracion.Logueado <> 3 Then
   Cliente.TCPSocket.Close
   SocketTCP.CambiarEstadoDelCliente (0) ' Poner el Estado Logueado
   ' Define que se termino el Loguin
   Me.MousePointer = vbDefault
   Exit Sub
 End If
 
 ' **************************************************************
 ' Graba la Password Actual en la Varible de Configuracion
 ' **************************************************************
 Configuracion.Password = Trim(Password)
 Configuracion.IDAliasUsuario = Trim(Usuario)
  
 ' **************************************************************
 ' Cambia el Estado Del Cliente
 ' **************************************************************
 SocketTCP.CambiarEstadoDelCliente (3) ' Poner el Estado Logueado
  
 ' **************************************************************
 ' Carga los Mensajes Pendientes
 ' **************************************************************
 Varios.CargarMensajesPendientes
 
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbDefault
 
 ' **************************************************************
 ' Graba el Usuario ya que se valido exitosamente
 ' **************************************************************
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Varios", "UltimoUsuario", Trim(Me.Usuario)
 
 ' **************************************************************
 ' Graba los Seteos Actuales
 ' **************************************************************
 If Me.RecordarPasswordCheck = 1 Then
   Configuracion.RecordarPasswordEstado = True
   Configuracion.RecordarPasswordPassword = Trim(Me.Password)
  Else
   Configuracion.RecordarPasswordEstado = False
   Configuracion.RecordarPasswordPassword = ""
 End If
 Inicializar.GrabarConfiguracion "NO"
 
 ' **************************************************************
 ' Carga la configuracion para traer los seteos propios del usuario
 ' **************************************************************
 Inicializar.CargarConfiguracionArchivo
  
 ' **************************************************************
 ' Cierra el Formulario de Loguin
 ' **************************************************************
 ' Graba el Estado
 Inicializar.UltimoEstado "Grabar", CStr(Usuario)
 '''
 Unload Loguin
 
End Sub
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
Private Sub EstadoVisible_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"

End Sub

Private Sub Form_Activate()

 ' **************************************************************
 ' Logueo Automatico
 ' **************************************************************
 If Configuracion.LogueoAutomatico And Variables.LoguinAutomatico And Me.ActivoAutologueo = False Then
  Me.ActivoAutologueo = True ' Esto lo hace para evitar que entre en Loop...
  BotonConectar
 End If
 
End Sub

Public Sub CargarTextos()

 ' **************************************************************
 ' Carga el Formulario segun el Lenguaje y los Colores...
 ' **************************************************************
 ' **************************************************************
 ' TituloVentana
 ' **************************************************************
 TituloVentana1.ForeColor = Variables.FontTituloVentana
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(117)
 ' **************************************************************
 ' Hypervinculos
 ' **************************************************************
 Me.RecordarPassword.ForeColor = Variables.FontHipervinculoColor
 Me.RecordarPassword = MensajeRecurso(101) ' Si Olvido su Password haga Click Aqui...
 ' **************************************************************
 ' Label's
 ' **************************************************************
 Me.UsuarioLbl.ForeColor = Variables.FontLabelColor
 Me.UsuarioLbl = MensajeRecurso(102) ' Usuario:
 Me.PasswordLbl.ForeColor = Variables.FontLabelColor
 Me.PasswordLbl = MensajeRecurso(103) ' Password:
 Me.Label4.ForeColor = Variables.FontLabelColor
 Me.Label4 = MensajeRecurso(104) ' ¿Disponible?
 Me.Label5.ForeColor = Variables.FontLabelColor
 Me.Label5 = MensajeRecurso(105) ' ¿Recordar Password?
 ' **************************************************************
 ' Botones
 ' **************************************************************
 Me.CancelarLbl = MensajeRecurso(106) ' Cancela
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Label2 = MensajeRecurso(107) ' Conectar...
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 ' **************************************************************
 ' Imaganes
 ' **************************************************************
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture


End Sub

Private Sub Form_Load()
Dim Usuario As String

 ' **************************************************************
 ' Define el Titulo de la Ventana
 ' **************************************************************
 ' Carga los Textos...
 Me.CargarTextos
 ' Nombre del Formulario...
 Me.FormularioNombre = "Loguin"
 ' Craga el Fondo del Formulario...
 'Me.Picture = LoadResPicture(Me.Name, "FONDO_FORMULARIOS")

 
 ' **************************************************************
 ' Define que aun no se disparo el autologueo
 ' **************************************************************
 Me.ActivoAutologueo = False
  
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 
 ' **************************************************************
 ' Verifica qaue no exista un Usuario ya cargado, si existe lo
 ' Carga...
 ' **************************************************************
 Usuario = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Varios", "UltimoUsuario")
 If Trim(Usuario) <> "" Then
  Me.Usuario = Usuario
 End If
 
 ' **************************************************************
 ' Carga el Autologueo
 ' **************************************************************
 If Configuracion.RecordarPasswordEstado Then
   Me.RecordarPasswordCheck = 1
   Me.Password = Configuracion.RecordarPasswordPassword
  Else
   Me.RecordarPasswordCheck = 0
   Me.Password = ""
 End If
 
 ' **************************************************************
 ' Logueo Automatico
 ' **************************************************************
 Variables.LoguinAutomatico = True
 
 ' **************************************************************
 ' Disponible o no
 ' **************************************************************
 If Trim(Usuario) <> "" Then
  Inicializar.UltimoEstado "leer", Usuario
 End If
 If Variables.UltimoEstadoDelUsuario.EstadoNumero = 2 Then
   Me.EstadoVisible = 0
  Else
   Me.EstadoVisible = 1
 End If
 
End Sub
Private Sub Form_Unload(Cancel As Integer)

 ' Ojo esto tiene que esta igual en BotonCancelar...
 ' **************************************************************
 ' Verifica si el Usuario esta logueado, si no es asi, cierra
 ' el sistema
 ' **************************************************************
  ' **************************************************************
 ' Define el Estado del Formulario
 ' **************************************************************
 Variables.FormularioLoguin = False

 If Configuracion.Logueado <> 3 Then
  Cliente.TCPSocket.Close
  SocketTCP.CambiarEstadoDelCliente (0)
 End If
 
    
End Sub
Private Sub Image3_Click() ' Boton Cancelar...
 
 ' **************************************************************
 ' Ejecuta el Sonido del Click...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 ' Cancelar...
 BotonCancelar
 
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

 ' Cancelar...
 BotonCancelar
 
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
 ' Ejecuta la Coneccion
 ' **************************************************************
 BotonConectar
 
End Sub
Private Sub Password_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Cuando Presiona Enter en Password Comienza la Validacion
 ' **************************************************************
 If KeyAscii = 13 Then
  BotonConectar
 End If
 
End Sub
Private Sub RecordarPassword_Click()
Dim SegundosTranscurridos As Integer
Dim PaqueteEnviar As String

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Verifica que se halla ingresado un Usuario
 ' **************************************************************
 If Trim(Usuario) = "" Then
  ' Debe Ingresar Un Usuario Valido...
  MostrarMSGBox MensajeRecurso(118), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Usuario.SetFocus
  Exit Sub
 End If
  
 ' **************************************************************
 ' Dispara el Timer de La Animacion
 ' **************************************************************
 Animacion.Enabled = True
 
 ' **************************************************************
 ' Comienza el Proceso de Envio de Password
 ' **************************************************************
 ' Esta variable es usada para determinar si el Servidor contesto
 ' que se envio o no la password, definiendo el estado:
 '      -1: No se contacto el Servidor...
 '       0: No se pudo enviar la Password...
 '       1: La password fue enviada a enviopassworddireccionmail
 '       2: El usuario no posee direccion de Email
 '       3: EL usuario no existe...
 EnvioPassword = -1
 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbHourglass
 ' Primero Cierra el Socket
 Cliente.TCPSocket.Close
 ' Comienza la Coneccion contra el Servidor
 Cliente.TCPSocket.RemoteHost = Trim(Configuracion.Servidor)
 Cliente.TCPSocket.RemotePort = Trim(Configuracion.PortTCP)
 Cliente.TCPSocket.Connect
 
 ' **************************************************************
 ' Espera 3 Segundos para que se conecte el Socket
 ' **************************************************************
 TiempoLogueoInicial = Time
 Do Until Cliente.TCPSocket.State = sckConnected
  DoEvents
  SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 ' Verifica que el Port se halla conectado... Sino es asi sale con errror...
 If Cliente.TCPSocket.State <> sckConnected Then
  ' No fue Posible Conectarse Contra el Servidor...
  MostrarMSGBox MensajeRecurso(108), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  ' Detiene la Animacion...
  Animacion.Enabled = False
  AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
  ' **
  SocketTCP.CambiarEstadoDelCliente (0)
  Me.MousePointer = vbDefault
  Exit Sub
 End If
 ' **************************************************************
 
 ' **************************************************************
 ' Enviar Paquete de Envio de Password por Mail
 ' **************************************************************
 PaqueteEnviar = "02" & CompletarCadena(CStr(Usuario), 16, "D", " ")
 EnviarPaqueteTCP (PaqueteEnviar)
  
 ' **************************************************************
 ' Traba los Controles mientras negocia el Logueo
 ' **************************************************************
 Usuario.Enabled = False
 Password.Enabled = False
 EstadoVisible.Enabled = False
 Label1.Enabled = False
 Label3.Enabled = False
 RecordarPassword.Enabled = False
 ' **************************************************************
 
 ' **************************************************************
 ' Deja un Bucle Para Valida si se envio la Passwor dpor Mail,
 ' tomando tiempo de TimeOut 5 segundos...
 TiempoLogueoInicial = Time
 Do
  DoEvents
  ' Si el servidor contesto el Envio de la Password sale
  If EnvioPassword <> -1 Then Exit Do
  
  ' Verifica que no se pase el Tiempo de TimeOut
  SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 ' **************************************************************
 
 
 ' **************************************************************
 ' DesTraba los Controles mientras negocia el Logueo
 ' **************************************************************
 Usuario.Enabled = True
 Password.Enabled = True
 EstadoVisible.Enabled = True
 Label1.Enabled = True
 Label3.Enabled = True
 RecordarPassword.Enabled = True
 ' **************************************************************
 
 ' **************************************************************
 ' Avisar Si Hubo algun Problema en el Logueo
 ' **************************************************************
 Select Case EnvioPassword
  Case -1:
   ' No fue Posible Conectarse Contra el Servidor...
   MostrarMSGBox MensajeRecurso(108), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Case 0:
   ' Se produjo un Error al Intentar Enviar su Password...
   MostrarMSGBox MensajeRecurso(119), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Case 1:
   ' La Password fue Enviada a [ % ]...
   MostrarMSGBox MensajeRecurso(120) & EnvioPasswordDireccionMail & MensajeRecurso(121), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Case 2:
   ' Su Usuario no Posee Direccion de E-Mail...
   MostrarMSGBox MensajeRecurso(122), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Case 3:
   ' El Usuario ' % ' No Existe...
   MostrarMSGBox MensajeRecurso(174) & Usuario & MensajeRecurso(124), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
 End Select
 
 ' Detiene la Animacion...
 Animacion.Enabled = False
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 ' **

 ' Cambia el Puntero del Mouse
 Me.MousePointer = vbDefault
 
End Sub
Private Sub RecordarPasswordCheck_Click()

 ' **************************************************************
 ' Cambia la configuracion de Recordar Password
 ' **************************************************************
 If Me.RecordarPasswordCheck = 1 Then
   Configuracion.RecordarPasswordEstado = True
  Else
   Configuracion.RecordarPasswordEstado = False
 End If
 
 ' **************************************************************
 ' Graba el Seteo
 ' **************************************************************
 GrabarConfiguracion "NO"
 
End Sub

Private Sub Usuario_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Cuando Presiona Enter, Automaticamente pasa al campo Password
 ' **************************************************************
 If KeyAscii = 13 Then
  Password.SetFocus
 End If

End Sub
