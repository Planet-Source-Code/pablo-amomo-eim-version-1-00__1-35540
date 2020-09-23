Attribute VB_Name = "Variables"
Option Explicit

' **************************************************************
' Maxima concurrencia en Multichat
' **************************************************************
Public Const MaximoMultichat = 8

' **************************************************************
' Volver a Preguntar para incluirlo
' **************************************************************
Public VolverPreguntarNombre() As String
Public VolverPreguntarCantidad As Integer

' **************************************************************
' Colores...
' **************************************************************
Public ShapesBackColor  As Long
Public ShapesBorderColor As Long
Public FontLabelColor As Long
Public FontHipervinculoColor As Long
Public FontTituloVentana As Long
Public FontBotonesColor As Long
Public FontMenuDescolgable As Long
Public FontFOndoMenuDescolgable As Long
Public FontMenuDescolgableAbierto As Long
Public FontMenuDescolgableAbiertoFranjas As Long
Public FontMenuDescolgableAbiertoFondoFormulario As Long
Public FontMenuDescolgableAbiertoLineaOscura As Long
Public FontMenuDescolgableAbiertoFont As Long
Public FontMenuDescolgableAbiertoHighLightFondo As Long
Public FontMenuDescolgableAbiertoHighLightLetra As Long
Public FontMenuDescolgableAbiertoMostrarGrafico As Boolean
Public FontMenuDescolgableAbiertoHighLightBorde As Long
Public FontFormClienteTextosDeEstado As Long
Public FontFormClienteTimer As Long
Public FontCambioDeEstado1 As Long
Public FontCambioDeEstado2 As Long
Public FontCambiarLetraNormalBack As Long
Public FontCambiarLetraNormalBorder As Long
Public FontCambiarLetraHighBack As Long
Public FontCambiarLetraHighBorder As Long
Public FontCambiarLetraLetra As Long
Public FontMensajeBotonApretadoFondo As Long
Public FontMensajeBotonColorLetra As Long

' **************************************************************
' Sonidos...
' **************************************************************
Public BufferSonido1() As Byte
Public BufferSonido2() As Byte
Public BufferSonido3() As Byte
Public BufferSonido4() As Byte
Public BufferSonido5() As Byte

Public LenguajeActual As String
Public PosicionVentana As POINTAPI

' **************************************************************
' Estas variables son utilizadas para definir si esta o no expandido un nodo...
' **************************************************************
Public GrupoEstadoNombre() As String
Public GrupoEstadoCantidad As Integer
Public GrupoEstadoPrimeraVez As Boolean

' **************************************************************
' variable para guardar el Ultimo Estado
' **************************************************************
Type MiUltimoEstado
 EstadoNumero As Integer
 Estadotexto As String
End Type
Public UltimoEstadoDelUsuario As MiUltimoEstado

' **************************************************************
' variable Utilizada para volver de "Enseguida Vuelvo"
' **************************************************************
Public PasoAEnseguidaVuelvo As Boolean

' **************************************************************
' Maneja el Aviso de los Mensajes Pendientes...
' **************************************************************
Public UltimoMensajePendienteHorario As Date
Public UltimoMensajePendienteAviso As Boolean

' **************************************************************
' Esta Variable se utiliza para pasar a "Enseguida Vuelvo", si
' esta mas de xx tiempo sin mandar nada...
' **************************************************************
Public UltimoMensajeEnviado As Date

' **************************************************************
' Define una Variable que Permita identificar el ID del Mensaje
' **************************************************************
Public IDMensaje As Integer

' **************************************************************
' Variable Armado de Mensajes
' **************************************************************
Public CantidadDePaquetesPartidos As Integer
Type ArmadoDePaqueteTMP
 UltimoPaqueteRecibido As Date
 PaquetesTotales As Integer
 UsuarioEmisor As String
 MultiChat As Boolean
 Handle As Long
 Datos(9) As String
 IDMensaje As String
End Type
Public ArmadoDePaquete() As ArmadoDePaqueteTMP

' **************************************************************
' Fuentes del Equipo
' **************************************************************
Public NombreDeFuentes() As String
Public CantidadDeFuentes As Integer

' **************************************************************
' Bloqueo de Usuarios...
' **************************************************************
Type MiUsuarioBloqueado
 NombreDelUsuario As String * 16
End Type
Public UsuarioBloqueadoNombres() As MiUsuarioBloqueado
Public UsuarioBloqueadoCantidad As Integer

' **************************************************************
' Estado del Socket
' **************************************************************
Public SocketTransmitiendo As Boolean

' **************************************************************
' Mensajes Pendientes
' **************************************************************
Type MiMensajesPendientesFormatoDefinido
 MensajeDe As String * 16
 Mensaje As String * 3000
 HoraYFecha As String * 100
End Type
Type MiMensajesPendientes
 MensajeDe As String
 Mensaje As String
 HoraYFecha As String
End Type
Type MiMensajesPendientesAgrupados
 MensajeDe As String
 Cantidad As Integer
End Type
Public MensajesPendientes() As MiMensajesPendientes
Public CantidadDeMensajesPendientes As Integer
Public MensajesPendientesAgrupados() As MiMensajesPendientesAgrupados
Public CantidadDeMensajesPendientesAgrupados As Integer

' **************************************************************
' Define el Workspace de Pantalla del Usuario
' **************************************************************
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
    Global Const SPI_GETWORKAREA As Long = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public AreaDeTrabajo As RECT

' **************************************************************
' Declaraciones Varias
' **************************************************************
Public Const GrisClaro = &HC0C0C0
Public Const Negro = &H0&

' **************************************************************
' Windows On-Top
' **************************************************************
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

' **************************************************************
' Para definir la Posicion del Mouse
' **************************************************************
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

' **************************************************************
' Variable para Definir el Listado de Amigos
' **************************************************************
Type MiGrupoAmigo
 NombreDelGrupo As String * 20
 IDNombreDelAmigo As String * 16
 EstadoDelAmigoEstado As Integer
 EstadoDelAmigoTexto As String * 20
 NombreDelAmigo As String * 50
 Sexo As String * 1
 Existe As Boolean
 DireccionEMail As String * 50
End Type
Public GrupoAmigo() As MiGrupoAmigo
Public CantidadGrupoAmigo As Integer

' **************************************************************
' Define el Sexo temporalmente...
' **************************************************************
Public SexoTemporal As String

' **************************************************************
' Variable para definir si el Formulario de Logueo esta
' Levantado o no...
' **************************************************************
Public FormularioLoguin As Boolean
Public LoguinAutomatico As Boolean

' **************************************************************
' Variable que guarda los Seteos Actuales del Cliente
' **************************************************************
Public Type MiConfiguracion
 ' Sexo
 Sexo As String
 ' Lenguaje
 Lenguaje As String
 ' Tiempo maxima para pasar automaticamente a Enseguida Vuelvo...
 AvisarMensajesPendientes As Integer
 ' Tiempo maxima para pasar automaticamente a Enseguida Vuelvo...
 TiempoParaPasaraInactivo As Integer
 ' Carga Minimizado
 CargarMinimizado As Boolean
 ' Loeguear Automatico
 LogueoAutomatico As Boolean
 ' Arranque con Windows
 ArranqueConWindows As Boolean
 ' FontEstandar para los Mensajes
 FontEstandarNombre As String
 FontEstandarTamano As Integer
 ' Datos del Usuario
 IDAliasUsuario As String * 16
 Password As String * 12
 ' Datos Generales
 RecordarPasswordEstado As Boolean
 RecordarPasswordPassword As String * 20
 NombreDelSistema As String * 20
 VersionDelSistema As String * 5
 TituloVentanas As String * 50
 ' Datos de Configuracion
 PortTCP As Integer
 Servidor As String * 20
 
 TimeOutLogueo As Integer    ' Tiempo maximo para esperar Coneccion (s)
 TimeOutGeneral As Integer   ' Tiempo Maximo para esperar las Transmiciones (s)
 TimeOutMultiChat As Integer ' Tiempo Maximo para esperar respuesta de Multichat (s)
 
 ' Datos de Estado
 Logueado As Integer    ' Valores:
                        '               0 - No Logueado
                        '               1 - Usuario Erroneo
                        '               2 - Password Erroneo
                        '               3 - Ok
 EstadoDelUsuario As Integer    ' Estado Actual: Aca se informa cual es el Estado del Usuario
                                ' 0. No Conectado
                                ' 1. Disponible
                                ' 2. No Disponible
                                ' 3. Custom
 MiNombreYApellido As String * 50 ' Mi nombre y Apellido
 EstadoActualTexto As String * 20
 TiempoDeRefrescoAmigos As Integer ' En Minutos
 ' Opciones Varias
 InformarCambiosDeEstado As Boolean
 SonidoActivado As Boolean
 DirectorioDownload As String * 255
 DirectorioUpload As String * 255
End Type
' Define la Variable
Public Configuracion As MiConfiguracion
' **************************************************************

' **************************************************************
' Cambio de Password
' Esta variable se usa para esperar que la password
' sea cambiada exitosamente...
' **************************************************************
Public CambioDePasswordOk As Boolean

' **************************************************************
' Usado Para COntabilizar el Tiempo en Linea
' **************************************************************
Public TiempoEnLineaContanteDesde

' **************************************************************
' Respuesta del MensajeBOX
' **************************************************************
Public RespuestaMensajeBox As Integer

' **************************************************************
' Respuesta del IngresoBox
' **************************************************************
Public RespuestaIngresoBox As String

' **************************************************************
' Cambio de Estado
' **************************************************************
Public Type MiNuevoEstadoUsuario
 Numero As Integer
 texto As String * 20
End Type
Public NuevoEstadoUsuario As MiNuevoEstadoUsuario

' **************************************************************
' Bandera Usado para ver el Amigo que se esta Verificando
' **************************************************************
Public RespuestaAmigoACrearAlias As String
Public RespuestaAmigoACrearResultado As Integer
Public RespuestaAmigoACrearEstado As Integer
Public RespuestaAmigoACrearEstadoTexto As String
Public RespuestaAmigoACrearSexo As String

' **************************************************************
' Bandera Usado para la Busqueda de Amigos
' **************************************************************
Public RespuestaBusquedaAmigos As Integer

