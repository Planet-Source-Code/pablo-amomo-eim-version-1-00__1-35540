Attribute VB_Name = "Inicializacion"
Option Explicit
Sub CambiarEstadoDelSistema(Estado As String)

 ' **************************************************************
 ' Se define la accion en caso de que se este Bajando o Subiendo
 ' el servicio
 ' **************************************************************
 Select Case Estado
  Case "Up":
   Configuracion.EstadoDelSistema = "Up"
   Server.BotonArranque.Caption = "Stop System"
   Server.EstadoLabel = "System Running..."
   Server.EstadoImagen.Picture = Server.EstadoSistema.ListImages("Levantado").Picture
   EscribirEvento "System Ready to Receive User's...(" & Configuracion.UsuariosSoportados & " Concurrent User's)", vbBlue
  Case "Down":
   Configuracion.EstadoDelSistema = "Down"
   Server.BotonArranque.Caption = "Start System"
   Server.EstadoLabel = "System Down..."
   Server.EstadoImagen.Picture = Server.EstadoSistema.ListImages("Detenido").Picture
   EscribirEvento "System Down...", vbRed
 End Select
 
 ' **************************************************************
 ' Pone la Cantidad de Usuarios en 0...
 ' **************************************************************
 Configuracion.UsuariosConectadosAlSistemas = 0
 Server.CantidadUsuariosLabel = Configuracion.UsuariosConectadosAlSistemas & " de " & Configuracion.UsuariosSoportados & "..."
 
End Sub
Sub DetenerSistema()
Dim Contador As Integer

 ' **********************************************************************
 ' Cierra la Coneccion de la Base de Datos...
 ' **********************************************************************
 BaseDeDatos.BajarBaseDeDatos
 
 ' **********************************************************************
 ' Cierra todos los Sockets del Sistema
 ' **********************************************************************
 For Contador = 1 To Configuracion.UsuariosSoportados
  If Sockets(Contador).EstadoDelPort <> 0 Then
   Server.TCPSocket(Contador).Close
   Unload Server.TCPSocket(Contador)
  End If
 Next
 
 ' **********************************************************************
 ' Cierra el Socket 0 - Socket Principal del Sistema
 ' **********************************************************************
 Server.TCPSocket(0).Close
 
 ' Define al Sistema en Estado Operativo y cambia los Captions
 CambiarEstadoDelSistema ("Down")
 
End Sub
Sub InicializarSistema()
Dim Respuesta As Boolean

 ' **********************************************************************
 ' Ultimo Port
 ' **********************************************************************
 'Variables.UltimoPort = 0
 
 ' **********************************************************************
 ' Carga la Configuracion Inicial del Sistema
 ' **********************************************************************
 CargarConfiguracionInicial
 
 ' **********************************************************************
 ' Aviso de Evento...
 ' **********************************************************************
 Logs.EscribirEvento Chr$(13) & _
 " *********************************************************************************" & Chr$(13) & _
 " * IMPORTANT:If you like, the Server Can Run like a NT Service, to   *" & Chr$(13) & _
 " * instance the Service Start de Server with the '/I' command Line.   *" & Chr$(13) & _
 " * To Unninstall the Service Run with '/D' Command Line...                 *" & Chr$(13) & _
 " *********************************************************************************", vbRed
  
 ' **********************************************************************
 ' Prepara y Setea la Base de Datos para Operar
 ' **********************************************************************
 Respuesta = AbrirBaseDeDatos
 ' No Arranca el Sistema... Lo deja en DOWN...
 If Respuesta = False Then
  CambiarEstadoDelSistema ("Down")
  Exit Sub
 End If
 
 ' **********************************************************************
 ' Carga en Memoria los Usuarios registrados en el Sistema
 ' **********************************************************************
 CargarUsuarios
 
 ' **********************************************************************
 ' Inicializar el Primer Socket y dejar Listo el Sistema
 ' **********************************************************************
 InicializarSocket
 
 ' **********************************************************************
 ' Define al Sistema en Estado Operativo y Cambia los Captions
 ' **********************************************************************
 CambiarEstadoDelSistema ("Up")
 
End Sub
Sub CargarConfiguracionInicial()
Dim Contador As Integer

 ' **********************************************************************
 ' Configuracion del Sistema
 ' **********************************************************************
 With Configuracion
  .NombreDelSistema = "EIM"
  .VersionDelSistema = "1.00"
  .TituloVentanas = "" ' Se define mas Abajo
  .EstadoDelSistema = "Down"
  .UsuariosConectadosAlSistemas = 0
  .CantidadDeUsuarios = 0
 End With
 Configuracion.TituloVentanas = Configuracion.NombreDelSistema & " Server " & _
                                "Version " & Configuracion.VersionDelSistema
 ' **********************************************************************
 
 ' **********************************************************************
 ' Carga los Datos de Configuracion del Archivo de Configuracion
 ' **********************************************************************
 CargarConfiguracionArchivo
 
 ' **********************************************************************
 ' Redefine la Cantidad de Ports Disponibles
 ' **********************************************************************
 ReDim Sockets(Configuracion.UsuariosSoportados)
 ' **********************************************************************
 
 ' **********************************************************************
 ' Limpiar las Variables de Port
 ' **********************************************************************
 For Contador = 1 To Configuracion.UsuariosSoportados
  With Sockets(Contador)
   .EstadoDelPort = 0
   .IDNumericoUsuario = 0
   .IDAliasUsuario = ""
  End With
 Next
 ' **********************************************************************
    
End Sub
