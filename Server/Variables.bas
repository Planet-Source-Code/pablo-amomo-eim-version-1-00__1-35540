Attribute VB_Name = "Variables"
Option Explicit

' **************************************************************
' Constantes usadas por el IconTray
' **************************************************************
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
' **************************************************************

' **************************************************************
' Ultimo Port / Hora...
' **************************************************************
'Public UltimoPort As Integer
'Public UltimoEnvio As Date

' **************************************************************
' Base de Datos
' **************************************************************
Public DataBase As ADODB.Connection
Public rsTablaUsuarios As ADODB.Recordset
Public rsTablaLogs As ADODB.Recordset
Public rsMensajeOffLine As ADODB.Recordset

' **************************************************************
' Configuracion del Sistema
' **************************************************************
Type MiArchivoConfiguracion
 ' Ubicacion y Datos de la Base de Datos
 UbicacionBaseDeDatos As String * 200
 NombreDeLaBaseDeDatos As String * 30
 ' Datos de Funcionalidad del Sistema
 UsuariosSoportados As String * 5   ' Usuarios Soportados por el Sistema
 PortTCP As Integer                 ' Port TCP sobre el que Funciona el Sistema
 DireccionIPSMTP As String * 15     ' Direccion IP del Servidor SMTP
 DireccionEMAILAdministrador As String * 50 ' Direccion EMail Del Administrador
 PermitirCrear As Boolean           ' Permitir Crear usuarios por el Cliente
End Type
Public ArchivoConfiguracion As MiArchivoConfiguracion
' **************************************************************

' **************************************************************
' Configuracion del Sistema
' **************************************************************
Type MiConfiguracion
 ' Datos Generales
 NombreDelSistema As String
 VersionDelSistema As String
 TituloVentanas As String
 UsuariosConectadosAlSistemas As Integer
 ' Ubicacion y Datos de la Base de Datos
 UbicacionBaseDeDatos As String * 200
 NombreDeLaBaseDeDatos As String * 30
 ' Datos de Funcionalidad del Sistema
 UsuariosSoportados As String * 5   ' Usuarios Soportados por el Sistema
 CantidadDeUsuarios As String       ' Cantidad de Usuarios en La Base de Datos
 PortTCP As Integer                 ' Port TCP sobre el que Funciona el Sistema
 DireccionIPSMTP As String * 15     ' Direccion IP del Servidor SMTP
 DireccionEMAILAdministrador As String * 50 ' Direccion EMail Del Administrador
 PermitirCrear As Boolean           ' Permitir Crear usuarios por el Cliente
 EstadoDelSistema As String         ' Estado Del Sistema Up o Down
End Type
' **
Public Configuracion As MiConfiguracion
' **************************************************************

' **************************************************************
' Variable de Estado de Sockets
' **************************************************************
Type MiSocket
 EstadoDelPort As Integer
 ' Estado de Usuario:
 '  0. Desconectado
 '  1. Logueando
 '  2. Logueado
 IDNumericoUsuario As Integer
 IDAliasUsuario As String * 16
End Type
' **
Public Sockets() As MiSocket
' **************************************************************

' **************************************************************
' Datos de Los Usuarios
' **************************************************************
Type MiUsuario
 IDNumericoUsuario As Integer
 IDAliasUsuario As String * 16
 Password As String * 12
 PortActual As Integer
 ' Port Actual
 ' 0. No esta Conectado
 ' x. Puerto donde se encuentra conectado
 EstadoActualNumero As Integer
 ' Estado Actual: Aca se informa cual es el Estado del Usuario
 ' 0. No Conectado
 ' 1. Visible Normal
 ' 2. No Disponible
 ' 3. Custom
 EstadoActualTexto As String * 20
 ' EstadoActualTexo: Define el Mensaje Custom en caso que
 ' EstadoActualNumero sea 3
 ' ******************************
 ' Datos Generales
 ' ******************************
 ApellidoYNombre As String * 50
 DireccionDeEmail As String * 50
 FechaDeNacimiento As String * 10
 Edad As String * 2
 Sexo As String * 1
 UbicacionGeografica As String * 20
 Intencion As String * 20
 Humor As String * 20
 Ocupacion As String * 20
 SigNo As String * 15
 EstadoCivil As String * 1
 Telefono As String * 50
 OtraInfo As String * 150
 ListadoDeAmigos As String
 MensajesOffline As String
 UsuarioBloqueado As Boolean
 UltimoLogueo As String
End Type
' **
Public Usuarios() As MiUsuario
' **************************************************************

