Attribute VB_Name = "SocketTCP"
Option Explicit
Function EnviarCambioDeEstado(Numero As Integer, texto As String)

 ' **************************************************************
 ' Completar a 20 el Texto
 ' **************************************************************
 texto = CompletarCadena(texto, 20, "D", " ")

 ' **************************************************************
 ' Envia el Nuevo Estado
 ' **************************************************************
 EnviarPaqueteTCP ("10" & Numero & texto)

End Function
Function RecibirPaqueteTCP(ByRef PaqueteRecibido As Variant) As Variant

 ' **************************************************************
 ' Modulo que preprosesa lo que se recibe por el Socket
 ' **************************************************************
 RecibirPaqueteTCP = PaqueteRecibido
 ' **************************************************************
 
End Function
Function EnviarPaqueteTCP(ByRef PaqueteAEnviar As Variant) As String
Dim TiempoLogueoInicial As Date
Dim SegundosTranscurridos As Integer
 
 ' **************************************************************
 ' Define que el Socket esta transmitiendo
 ' **************************************************************
 Variables.SocketTransmitiendo = True

 ' **************************************************************
 ' Verificar el Estado del Port, si no esta conectado descarta
 ' el paquete....
 ' **************************************************************
 If Cliente.TCPSocket.State <> sckConnected Then
  If Cliente.TCPSocket.State = sckClosed Or Cliente.TCPSocket.State = sckError Then
   Cliente.TCPSocket.Close
   SocketTCP.CambiarEstadoDelCliente (0)
  End If
  Variables.SocketTransmitiendo = False
  Exit Function
 End If
 
 ' **************************************************************
 ' Modulo que preprosesa lo que se desea Enviar
 ' **************************************************************
 Cliente.TCPSocket.SendData (PaqueteAEnviar)
 ' **************************************************************
 
End Function
Function CambiarEstadoDelCliente(Estado As Integer)
Dim EstadoCLienteImagen, EstadoClienteTexto As String
Dim EstadoAnterior As Integer

 ' **************************************************************
 ' Guarda el Estado Anterior
 ' **************************************************************
 EstadoAnterior = Configuracion.Logueado
 
 ' **************************************************************
 ' En este modulo se cambia y centralizan todas las tareas cuando
 ' se cambia de Estado al Cliente (Logueado - No Logueado)
 ' **************************************************************
 ' Los estado son:
 '                  0 - No Logueado
 '                  1 - Conectando
 '                  3 - Logueado
 Configuracion.Logueado = Estado
 
 ' **************************************************************
 ' Esconde o Muestra los Menus segun el Estado...
 ' **************************************************************
 Select Case Estado
  Case 0:
   ' **************************************************************
   ' Baja todas las Ventanas...
   ' **************************************************************
   DescargarTodosLosFormularios
   
   ' **************************************************************
   ' Cambia el Icon Tray
   ' **************************************************************
   Varios.CambiarIconoTray "Desconectado"
   
   ' **************************************************************
   ' Timer Refresco de Amigos / Tiempo OnLine
   ' **************************************************************
   Cliente.RefrescoAmigos.Enabled = False
   Cliente.OnLineTime.Enabled = False
   
   ' **************************************************************
   ' Mensaje Pendientes...
   ' **************************************************************
   CambiarMensajesPendientes (0)
   
   ' **************************************************************
   ' Menu Coneccion
   ' **************************************************************
   ' Conectar = Si
   Cliente.MenuToolConeccion.HabilitarItem 0, True
   ' Desconectar = No
   Cliente.MenuToolConeccion.HabilitarItem 1, False
   ' Salir = Si
   Cliente.MenuToolConeccion.HabilitarItem 3, True
   ' **************************************************************
   ' Menu Configuracion
   ' **************************************************************
   ' Ver Mis Datos = No
   Cliente.MenuToolConfiguracion.HabilitarItem 0, False
   ' Cambiar Mis Datos = No
   Cliente.MenuToolConfiguracion.HabilitarItem 1, False
   ' Cambiar Password = No
   Cliente.MenuToolConfiguracion.HabilitarItem 3, False
   ' Preferencias = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 5, True
   ' **************************************************************
   ' Menu Amigos
   ' **************************************************************
   ' Recargar Amigos = No
   Cliente.MenuToolAmigos.HabilitarItem 0, False
   ' Crear Grup = No
   Cliente.MenuToolAmigos.HabilitarItem 2, False
   ' Borrar Grupo = No
   Cliente.MenuToolAmigos.HabilitarItem 3, False
   ' Agregar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 5, False
   ' Eliminar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 6, False
   ' Buscar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 7, False
   ' Bloquear Amigos = No
   Cliente.MenuToolAmigos.HabilitarItem 9, False
   
   ' **************************************************************
   ' Listado de Amigos
   ' **************************************************************
   ' Lo deja Abierto para poder ver los Mensajes Pendientes...
   ' Cliente.ListadoDeAmigos.Enabled = False
   Varios.CambiaEstadoListadoDeAmigos (0)
   ' **************************************************************
   
   ' **************************************************************
   ' Cambia el Estado de los MensajesPendiente
   ' **************************************************************
   Varios.CambiarMensajesPendientes Variables.CantidadDeMensajesPendientes
   
   ' **************************************************************
   ' Toolbars
   ' **************************************************************
   Cliente.ImageMenuDeEstado.Enabled = False
   
   ' **************************************************************
   ' Pasa el Estado del Usuario a Visible -
   ' **************************************************************
   Configuracion.EstadoDelUsuario = 1
   Configuracion.EstadoActualTexto = ""
   ' Define la Imagen y El Texto del Estado del Cliente
   EstadoCLienteImagen = "Desconectado"
   ' Desconectado...
   EstadoClienteTexto = MensajeRecurso(228) & "..."
   
   ' Define el Titulo de Ventana...
   'Configuracion.TituloVentanas = Trim(Trim(Configuracion.NombreDelSistema) & " " & Trim(Configuracion.VersionDelSistema))
   'CambiarTituloVentana
   CambiarTituloApp Configuracion.TituloVentanas & " - " & MensajeRecurso(228) & "..."
  Case 1: ' Conectando
   ' **************************************************************
   ' Cambia el Icon Tray
   ' **************************************************************
   Varios.CambiarIconoTray "Conectando"
   
   ' **************************************************************
   ' Timer Refresco de Amigos / Tiempo OnLine
   ' **************************************************************
   Cliente.RefrescoAmigos.Enabled = False
   Cliente.OnLineTime.Enabled = False
   
   ' **************************************************************
   ' Menu Coneccion
   ' **************************************************************
   ' Conectar = Si
   Cliente.MenuToolConeccion.HabilitarItem 0, True
   ' Desconectar = No
   Cliente.MenuToolConeccion.HabilitarItem 1, False
   ' Salir = Si
   Cliente.MenuToolConeccion.HabilitarItem 3, True
   ' **************************************************************
   ' Menu Configuracion
   ' **************************************************************
   ' Ver Mis Datos = No
   Cliente.MenuToolConfiguracion.HabilitarItem 0, False
   ' Cambiar Mis Datos = No
   Cliente.MenuToolConfiguracion.HabilitarItem 1, False
   ' Cambiar Password = No
   Cliente.MenuToolConfiguracion.HabilitarItem 3, False
   ' Preferencias = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 5, True
   ' **************************************************************
   ' Menu Amigos
   ' **************************************************************
   ' Recargar Amigos = No
   Cliente.MenuToolAmigos.HabilitarItem 0, False
   ' Crear Grup = No
   Cliente.MenuToolAmigos.HabilitarItem 2, False
   ' Borrar Grupo = No
   Cliente.MenuToolAmigos.HabilitarItem 3, False
   ' Agregar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 5, False
   ' Eliminar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 6, False
   ' Buscar Amigo = No
   Cliente.MenuToolAmigos.HabilitarItem 7, False
   ' Bloquear Amigos = No
   Cliente.MenuToolAmigos.HabilitarItem 9, False
   
   ' **************************************************************
   ' Listado de Amigos
   ' **************************************************************
   ' Lo deja Abierto para poder ver los Mensajes Pendientes...
   ' Cliente.ListadoDeAmigos.Enabled = False
   Varios.CambiaEstadoListadoDeAmigos (1)
   
   ' **************************************************************
   ' Cambia el Estado de los MensajesPendiente
   ' **************************************************************
   Varios.CambiarMensajesPendientes Variables.CantidadDeMensajesPendientes
   
   ' **************************************************************
   ' Toolbar
   ' **************************************************************
   Cliente.ImageMenuDeEstado.Enabled = False
   
   ' **************************************************************
   ' Pasa el Estado del Usuario a Visible -
   ' **************************************************************
   Configuracion.EstadoDelUsuario = 1
   Configuracion.EstadoActualTexto = ""
   ' Define la Imagen y El Texto del Estado del Cliente
   EstadoCLienteImagen = "Conectando"
   ' Conectando...
   EstadoClienteTexto = MensajeRecurso(229) & "..."
  
   ' Define el Titulo de Ventana...
   'Configuracion.TituloVentanas = Trim(Trim(Configuracion.NombreDelSistema) & " " & Trim(Configuracion.VersionDelSistema))
   'CambiarTituloVentana
   CambiarTituloApp Configuracion.TituloVentanas & " - " & MensajeRecurso(229) & "..."
  Case 3:
   ' **************************************************************
   ' Cambia el Icon Tray
   ' **************************************************************
   Varios.CambiarIconoTray "ConectadoSinMensaje"
   
   ' **************************************************************
   ' Cambia el Estado de los MensajesPendiente
   ' **************************************************************
   ' No hace Falta ya que cuando hace la recarga de Amigos cambia
   ' el Estado...
   ' Varios.CambiarMensajesPendientes 0
   
   ' **************************************************************
   ' Timers
   ' **************************************************************
   Cliente.RefrescoAmigos.Enabled = True
   Cliente.OnLineTime.Enabled = True
   
   ' **************************************************************
   ' Carga el estado de los Grupos
   ' **************************************************************
   Variables.GrupoEstadoPrimeraVez = True
   Inicializar.CargarEstadoDeGrupos

   ' **************************************************************
   ' Recarga el Listado de Amigos
   ' **************************************************************
   Varios.RecargarListadoDeAmigos
   
   ' **************************************************************
   ' Menu Coneccion
   ' **************************************************************
   ' Conectar = Si
   Cliente.MenuToolConeccion.HabilitarItem 0, False
   ' Desconectar = Si
   Cliente.MenuToolConeccion.HabilitarItem 1, True
   ' Salir = Si
   Cliente.MenuToolConeccion.HabilitarItem 3, True
   ' **************************************************************
   ' Menu Configuracion
   ' **************************************************************
   ' Ver Mis Datos = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 0, True
   ' Cambiar Mis Datos = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 1, True
   ' Cambiar Password = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 3, True
   ' Preferencias = Si
   Cliente.MenuToolConfiguracion.HabilitarItem 5, True
   ' **************************************************************
   ' Menu Amigos
   ' **************************************************************
   ' Recargar Amigos = Si
   Cliente.MenuToolAmigos.HabilitarItem 0, True
   ' Crear Grup = Si
   Cliente.MenuToolAmigos.HabilitarItem 2, True
   ' Borrar Grupo = Si
   Cliente.MenuToolAmigos.HabilitarItem 3, True
   ' Agregar Amigo = Si
   Cliente.MenuToolAmigos.HabilitarItem 5, True
   ' Eliminar Amigo = Si
   Cliente.MenuToolAmigos.HabilitarItem 6, True
   ' Buscar Amigo = Si
   Cliente.MenuToolAmigos.HabilitarItem 7, True
   ' Bloquear Amigos = Si
   Cliente.MenuToolAmigos.HabilitarItem 9, True
   
   ' **************************************************************
   ' Listado de Amigos
   ' **************************************************************
   Cliente.ListadoDeAmigos.Enabled = True
   
   ' **************************************************************
   ' Toolbar
   ' **************************************************************
   Cliente.ImageMenuDeEstado.Enabled = True
   
   ' **************************************************************
   ' Define la Imagen y El Texto del Estado del Cliente
   ' **************************************************************
   'If UCase(Trim(Configuracion.Sexo)) = "F" Then
   '  EstadoCLienteImagen = "Mujer"
   ' Else
   '  EstadoCLienteImagen = "Hombre"
   'End If
   EstadoCLienteImagen = "Conectado" 'aca
   ' Conectado...
   EstadoClienteTexto = MensajeRecurso(230) & "..."
      
   ' **************************************************************
   ' Define que el ultimo mensaje se mando cuando se conecto...
   ' **************************************************************
   Variables.UltimoMensajeEnviado = Time
   
   ' Define el Titulo de Ventana...
   'Configuracion.TituloVentanas = Trim(Trim(Configuracion.NombreDelSistema) & " " & Trim(Configuracion.VersionDelSistema)) & " - " & Trim(Configuracion.IDAliasUsuario)
   'CambiarTituloVentana
   CambiarTituloApp Configuracion.TituloVentanas & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
 End Select
 
 If EstadoAnterior <> 3 Then
  ' Pone a 0 el Contador de Tiempo en Linea
  Cliente.TiempoEnLinea = "00:00:00"
  Variables.TiempoEnLineaContanteDesde = Date & " " & Time
 End If
 
 ' **************************************************************
 ' Busca el Que sea datos de Usuario y Sea el Alias propio...
 ' **************************************************************
 Dim Contador As Integer
 For Contador = 1 To Forms.Count - 1
  Dim FormularioNombre, AliasUsuario As String
  ' **************************************************************
  ' Verifica que el Formulario sea de Datos
  ' **************************************************************
  If Forms(Contador).FormularioNombre = "DatosUsuario" Then
   FormularioNombre = Trim(Forms(Contador).FormularioNombre)
   AliasUsuario = Trim(Forms(Contador).AliasUsuario)
   If Trim(UCase(AliasUsuario)) = Trim(UCase(Configuracion.IDAliasUsuario)) Then
    Forms(Contador).PonerElEstadoDelUsuario
   End If
  End If
 Next
 ' **************************************************************

 ' Pone los Controles de Estado (Imagen y Texto)
 Cliente.EstadoCLienteImagen.Picture = Cliente.Imagenes.ListImages(EstadoCLienteImagen).Picture
 Cliente.EstadoClienteTexto = EstadoClienteTexto
 ' Si esta conectado agrega a EstadoClienteTexto el Estado del Cliente (On-Line, No Disponible, Custom)
 If Estado = 3 Then
   Select Case Configuracion.EstadoDelUsuario
    Case 1:
     'Cliente.Toolbar.Buttons(1).Image = Cliente.Imagenes.ListImages("EstadoVisible").Index
     Cliente.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoVisible").Picture
     ' Disponible (Normal)...
     Cliente.EstadoUsuarioTexto = " - " & MensajeRecurso(180)
    Case 2:
     'Cliente.Toolbar.Buttons(1).Image = Cliente.Imagenes.ListImages("EstadoNoDisponible").Index
     Cliente.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoNoDisponible").Picture
     ' No Disponible...
     Cliente.EstadoUsuarioTexto = " - " & MensajeRecurso(181)
    Case 3:
     'Cliente.Toolbar.Buttons(1).Image = Cliente.Imagenes.ListImages("EstadoCustom").Index
     Cliente.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoCustom").Picture
     If Len(Configuracion.EstadoActualTexto) > 20 Then
       Cliente.EstadoUsuarioTexto = " - " & Trim(Mid$(Configuracion.EstadoActualTexto, 1, 20)) & "..."
      Else
       Cliente.EstadoUsuarioTexto = " - " & Trim(Configuracion.EstadoActualTexto)
     End If
   End Select
  Else
   ' Desconectado...
   Cliente.EstadoUsuarioTexto = " - " & MensajeRecurso(228) & "..."
   Cliente.EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("NoConectado").Picture
 End If
 
 
End Function
Public Function CambiarTituloApp(texto As String)

 App.Title = texto
 
End Function

'Public Function CambiarTituloVentana()
'Dim Contador As Integer
'Dim Nombre As String
'
' ' ***************************************************************
' ' Pone el Titulo en todas las ventanas existentes...
' ' ***************************************************************
' For Contador = 0 To Forms.Count - 1
'  Nombre = UCase(Trim(Forms(Contador).FormularioNombre))
'  If Nombre <> "PRESENTACION" And Nombre <> "VENTANAMENU" And Nombre <> "CAMBIODEESTADO" Then
'   Forms(Contador).TituloVentana1 = Configuracion.TituloVentanas
'  End If
' Next
'End Function
