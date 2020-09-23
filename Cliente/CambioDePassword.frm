VERSION 5.00
Begin VB.Form CambioDePassword 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "CambioDePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CambioDePassword.frx":000C
   ScaleHeight     =   2400
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Animacion 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1920
      Top             =   1830
   End
   Begin VB.TextBox PasswordNuevaRepetir 
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1290
      Width           =   1635
   End
   Begin VB.TextBox PasswordNueva 
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   920
      Width           =   1665
   End
   Begin VB.TextBox PasswordActual 
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   530
      Width           =   1665
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4030
      MouseIcon       =   "CambioDePassword.frx":2559
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4005
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
      TabIndex        =   11
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   3960
      Top             =   1320
      Width           =   270
   End
   Begin VB.Label RepetirPWDLBL 
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
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1290
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2520
      MouseIcon       =   "CambioDePassword.frx":26AB
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1830
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
      TabIndex        =   7
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "CambioDePassword.frx":27FD
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1830
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
   Begin VB.Label NuevaPasswordLbl 
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   930
      Width           =   1755
   End
   Begin VB.Label PWDActualLBL 
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
      Left            =   240
      TabIndex        =   0
      Top             =   570
      Width           =   1785
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
Attribute VB_Name = "CambioDePassword"
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
' Esta variable guarda la fecha/hora para calcular el tiempo de Cambio de PWD
Private TiempoLogueoInicial As Date
Private IndiceAnimacion As Integer
Sub BotonCancelar()
 
 ' **************************************************************
 ' Descargar el Formulario
 ' **************************************************************
 Unload Me
 
End Sub
Private Sub BotonCambiarPassword()
Dim SegundosTranscurridos As Integer
Dim PaqueteEnviar, PWDTemp, PWDTemp2 As String

 ' **************************************************************
 ' Verifica que el Socket este abierto y si esta Logueado
 ' **************************************************************
 If Cliente.TCPSocket.State <> sckConnected Or Configuracion.Logueado = 0 Then
  'Unload Me
  BotonCancelar
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica que la Password Actual Sea Correcta
 ' **************************************************************
 If Trim(PasswordActual) <> Trim(Configuracion.Password) Then
  ' Muestra: La Password Actual Ingresada No es Valida...
  MostrarMSGBox MensajeRecurso(155), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  PasswordActual.SetFocus
  Exit Sub
 End If

 ' **************************************************************
 ' Le saca los Blancos a la Password Nueva...
 ' **************************************************************
 PasswordNueva = Trim(PasswordNueva)
 PasswordNuevaRepetir = Trim(PasswordNuevaRepetir)
 
 ' **************************************************************
 ' Verifica que la Nueva Password no Sea Nula
 ' **************************************************************
 If PasswordNueva = "" Then
  ' La Nueva Password no puede ser Nula...
  MostrarMSGBox MensajeRecurso(156), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  PasswordNueva.SetFocus
  Exit Sub
 End If
 
 ' **************************************************************
 ' Primero Verifica que la PWD Nueva y su Verificacion
 ' Sean iguales
 ' **************************************************************
 If PasswordNueva <> PasswordNuevaRepetir Then
  ' Muestra: La Nueva Password y su Verificación no Coinciden...
  MostrarMSGBox MensajeRecurso(157), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  PasswordNuevaRepetir.SetFocus ' Se posiciona en la Password Nueva Repetir...
  Exit Sub
 End If
 
 ' **************************************************************
 ' Define la Variable de Espera
 ' **************************************************************
 CambioDePasswordOk = False
 
 ' **************************************************************
 ' Dispara la Animacion
 ' ************************************************************
 BloqueaDesbloqueaControles False
 'Animacion.Enabled = True
 'PasswordActual.Enabled = False
 'PasswordNueva.Enabled = False
 'PasswordNuevaRepetir.Enabled = False
 'Label1.Enabled = False
 'Label3.Enabled = False
 ' Cambia el Puntero del Mouse
 'Me.MousePointer = vbHourglass
 ' **************************************************************
  
 ' **************************************************************
 ' Comienza el Proceso de Cambio de PWD
 ' ************************************************************
 ' Encripar la Password Nueva con la Password Actual
 PWDTemp = Encriptar(PasswordNueva, Configuracion.Password)
 ' Encripta la Password Actual Encriptada con ella misma
 PWDTemp2 = Encriptar(PasswordActual, Configuracion.Password)
 ' Arma el Paquete de Loguin
 'PaqueteEnviar = "01" & _
                 CompletarCadena(CStr(PWDTemp), 12, "I", " ") & _
                 CompletarCadena(CStr(PWDTemp2), 12, "I", " ")
 ' Enviar Paquete de Cambio de Password
 'EnviarPaqueteTCP (PaqueteEnviar)
 EnviarPaqueteTCP "01" & _
                  CompletarCadena(CStr(PWDTemp), 12, "I", " ") & _
                  CompletarCadena(CStr(PWDTemp2), 12, "I", " ")
  
 ' **************************************************************
 ' Deja un Bucle Para esperar la Confirmacion del Cambio
 ' de PWD
 ' **************************************************************
 TiempoLogueoInicial = Time
 Do
  DoEvents
  ' Si el Cambio de PWD es exitoso sale del Bucle
  If CambioDePasswordOk = True Then Exit Do
  ' Verifica que no se pase el Tiempo de TimeOut
  'SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
  'If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
  If DateDiff("s", TiempoLogueoInicial, Time) >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 
 ' **************************************************************
 ' Destraba los Controles
 ' **************************************************************
 BloqueaDesbloqueaControles True
 ' Cambia el Puntero del Mouse
 'Me.MousePointer = vbDefault
 'PasswordActual.Enabled = True
 'PasswordNueva.Enabled = True
 'PasswordNuevaRepetir.Enabled = True
 'Label1.Enabled = True
 'Label3.Enabled = True
 'Animacion.Enabled = False
 'AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 ' **************************************************************
 
 ' **************************************************************
 ' Avisar Si No se pudo Cambiar la Password
 ' **************************************************************
 If CambioDePasswordOk = False Then
  ' Muestra: No Fue Posible Cambiar la Password...
  MostrarMSGBox MensajeRecurso(158), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub ' Sale !
 End If
 
 ' **************************************************************
 ' Graba la Nueva Password
 ' **************************************************************
 Configuracion.Password = PasswordNueva
 
 ' **************************************************************
 ' Cambio de Password Exitoso...
 ' **************************************************************
 ' Muestra: El Cambio de Password Fue Realizado Exitosamente...
 MostrarMSGBox MensajeRecurso(159), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 
 ' **************************************************************
 ' Cierra el Formulario de Cambio de Password
 ' **************************************************************
 BotonCancelar
 'Unload CambioDePassword
 
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
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(IndiceAnimacion).Picture
  
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
Public Sub CargarTextos()

 ' **************************************************************
 ' Define el Texto del Formulario...
 ' **************************************************************
 ' Titulo Ventana
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 TituloVentana1 = Trim(Configuracion.TituloVentanas) & MensajeRecurso(160)
 ' Lables...
 Me.PWDActualLBL.ForeColor = Variables.FontLabelColor
 Me.PWDActualLBL = MensajeRecurso(150) ' Password Actual:
 Me.NuevaPasswordLbl.ForeColor = Variables.FontLabelColor
 Me.NuevaPasswordLbl = MensajeRecurso(151) ' Nueva Password:
 Me.RepetirPWDLBL.ForeColor = Variables.FontLabelColor
 Me.RepetirPWDLBL = MensajeRecurso(152) ' Repetir Password:
 ' Botones...
 Me.CancelarBt.BackColor = Variables.ShapesBackColor
 Me.CancelarBt.BorderColor = Variables.ShapesBorderColor
 Me.CancelarLbl.ForeColor = Variables.FontBotonesColor
 Me.CancelarLbl = MensajeRecurso(106) ' Cacenlar
 Me.Shape3.BackColor = Variables.ShapesBackColor
 Me.Shape3.BorderColor = Variables.ShapesBorderColor
 Me.Label2.ForeColor = Variables.FontBotonesColor
 Me.Label2 = MensajeRecurso(154) ' Cambiar Password
 ' Imagenes...
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture

End Sub
Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos...
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon

End Sub
Private Sub Form_Unload(Cancel As Integer)

 ' **************************************************************
 ' Aca llama al Boton Cancelar, para procesar la misma accion...
 ' **************************************************************
 'BotonCancelar
  
End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Cancela !
 ' **************************************************************
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

 ' **************************************************************
 ' Cancela !
 ' **************************************************************
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
 ' Cambiar Password...
 ' **************************************************************
 BotonCambiarPassword
 
End Sub
Private Sub PasswordActual_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Cuando presiona Enter Pasa al Nuevo Campo
 ' **************************************************************
 If KeyAscii = 13 Then
  PasswordNueva.SetFocus
 End If
 
End Sub
Private Sub PasswordNueva_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Cuando presiona Enter Pasa al Nuevo Campo
 ' **************************************************************
 If KeyAscii = 13 Then
  PasswordNuevaRepetir.SetFocus
 End If

End Sub
Private Sub PasswordNuevaRepetir_KeyPress(KeyAscii As Integer)

 ' **************************************************************
 ' Cuando presiona Enter Procesa el Cambio de PWD
 ' **************************************************************
 If KeyAscii = 13 Then
  BotonCambiarPassword
 End If

End Sub
Private Sub BloqueaDesbloqueaControles(Estado As Boolean)

 ' **************************************************************
 ' Dispara la Animacion
 ' ************************************************************
 Animacion.Enabled = Not Estado
 PasswordActual.Enabled = Estado
 PasswordNueva.Enabled = Estado
 PasswordNuevaRepetir.Enabled = Estado
 Label1.Enabled = Estado
 Label3.Enabled = Estado
 If Estado Then
   Me.MousePointer = vbHourglass
   AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
  Else
   Me.MousePointer = vbDefault
 End If
 ' **************************************************************

End Sub

