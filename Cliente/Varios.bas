Attribute VB_Name = "Varios"
Option Explicit
Public Function VolverAPreguntar(Usuario As String, Comando As String) As Integer
Dim Contador As Integer
Dim Bandera As Boolean

 Select Case UCase(Comando)
  Case "AGREGAR"
   ' Verificar si Existe...
   Bandera = False
   For Contador = 1 To VolverPreguntarCantidad
    If UCase(Trim(VolverPreguntarNombre(Contador))) = UCase(Trim(Usuario)) Then
     Bandera = True
     Exit For
    End If
   Next
   ' Lo Agrega....
   If Bandera = False Then
    VolverPreguntarCantidad = VolverPreguntarCantidad + 1
    ReDim Preserve Variables.VolverPreguntarNombre(VolverPreguntarCantidad)
    Variables.VolverPreguntarNombre(VolverPreguntarCantidad) = Usuario
   End If
  Case "BUSCAR"
   ' Lo Busca
   Bandera = False
   For Contador = 1 To VolverPreguntarCantidad
    If UCase(Trim(VolverPreguntarNombre(Contador))) = UCase(Trim(Usuario)) Then
     Bandera = True
     Exit For
    End If
   Next
   VolverAPreguntar = Bandera
 End Select
 
End Function

Public Function CrearNuevoAmigo(Usuarioexiste As Boolean, NombreGrupo As String, AmigoACrear As String, Estadoamigo As Integer, EstadoAmigoTexto As String, Sexo As String)
Dim TiempoInicial As Date

 ' **************************************************************
 ' Crea el Nuevo Amigo
 ' **************************************************************
 Variables.CantidadGrupoAmigo = Variables.CantidadGrupoAmigo + 1
 ReDim Preserve Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
 With Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
  .Existe = Trim(Usuarioexiste)
  .NombreDelGrupo = Trim(NombreGrupo)
  .NombreDelAmigo = ""
  .IDNombreDelAmigo = Trim(AmigoACrear)
  .Sexo = UCase(Trim(Sexo))
  .EstadoDelAmigoEstado = Estadoamigo
  .EstadoDelAmigoTexto = Trim(EstadoAmigoTexto)
 End With
  
 ' **************************************************************
 ' Grabar la Eliminacion
 ' **************************************************************
 Varios.GrabarListadoAmigos
 
 ' **************************************************************
 ' Espera un Segundo...
 ' **************************************************************
 TiempoInicial = Time
 Do Until DateDiff("s", TiempoInicial, Time) >= 1
  DoEvents
 Loop
  
 ' **************************************************************
 ' Avisa al Amigo que lo agrego...
 ' **************************************************************
 'EnviarPaqueteTCP "3" & CompletarCadena(Trim(AmigoACrear), 16, "D", " ") & "6" & CompletarCadena(Trim(Configuracion.MiNombreYApellido), 50, "D", " ")
 Dim Mensaje As String
 Mensaje = Chr$(0) & Chr$(0) & CompletarCadena(Trim(Configuracion.IDAliasUsuario), 16, "D", " ") & Trim(Configuracion.MiNombreYApellido)
 'Mensaje = MensajeRecursoReal(465) & " " & Trim(Configuracion.MiNombreYApellido) & " " & MensajeRecursoReal(466) & vbCrLf
 'Mensaje = Mensaje & MensajeRecursoReal(1465) & " " & Trim(Configuracion.MiNombreYApellido) & " " & MensajeRecursoReal(1466)
 EnviarPaqueteTCP "5" & CompletarCadena(CStr(Trim(AmigoACrear)), 16, "D", " ") & Mensaje
 
End Function
Public Function ArreglarAmigosRepetidos() As Boolean
Dim GrupoAmigoTemporal() As MiGrupoAmigo
Dim Borrar() As Boolean
Dim Contador, Contador2, Cantidad As Integer
Dim Buscando As String

 ' **************************************************************
 ' Primero Carga Todo en una variable Temporal...
 ' **************************************************************
 ReDim GrupoAmigoTemporal(Variables.CantidadGrupoAmigo)
 ReDim Borrar(Variables.CantidadGrupoAmigo)
 For Contador = 1 To Variables.CantidadGrupoAmigo
  GrupoAmigoTemporal(Contador) = GrupoAmigo(Contador)
  Borrar(Contador) = False
 Next
 
 ' **************************************************************
 ' Busca los Duplicados en Temporal, y si encuentra mas de un
 ' usuario ID, a partir del 2do. le pone en la vaiable borrar true
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  Buscando = Trim(GrupoAmigo(Contador).IDNombreDelAmigo)
  Borrar(Contador) = False
  Cantidad = 0
  If Trim(Buscando) <> "" Then  ' Si el Nombre es "" es por que es un Grupo...
                                ' Por tando no hace nada (No lo borra ni nada)
   For Contador2 = 1 To Variables.CantidadGrupoAmigo
    If UCase(Trim(GrupoAmigoTemporal(Contador2).IDNombreDelAmigo)) = UCase(Trim(Buscando)) Then
     Cantidad = Cantidad + 1
     If Cantidad >= 2 Then
      GrupoAmigoTemporal(Contador).IDNombreDelAmigo = ""
      Borrar(Contador) = True
     End If
    End If
   Next
  End If
 Next
 
 ' **************************************************************
 ' Carga los Nuevos en GrupoAmigo
 ' **************************************************************
 Cantidad = 0
 For Contador = 1 To Variables.CantidadGrupoAmigo
  'If Trim(GrupoAmigoTemporal(Contador).IDNombreDelAmigo) <> "" Then
  If Borrar(Contador) = False Then
   Cantidad = Cantidad + 1
   GrupoAmigo(Cantidad) = GrupoAmigoTemporal(Contador)
  End If
 Next
 If Variables.CantidadGrupoAmigo <> Cantidad Then
   Variables.CantidadGrupoAmigo = Cantidad
   ArreglarAmigosRepetidos = True
  Else
   ArreglarAmigosRepetidos = False
 End If
 
End Function
Public Function DescargarTodosLosFormularios()
Dim Contador

 ' **************************************************************
 ' Descarga Todos los Formularios...
 ' **************************************************************
 For Contador = Forms.Count - 1 To 0 Step -1
  If Forms(Contador).FormularioNombre <> "Cliente" And Forms(Contador).FormularioNombre <> "Loguin" And Forms(Contador).FormularioNombre <> "VentanaMenu" Then 'And Forms(Contador).FormularioNombre <> "Mensajes" Then
   If Forms(Contador).Visible = True Then
    Unload Forms(Contador)
    'Forms(Contador).Hide
   End If
  End If
 Next
 

End Function
Public Function VerificarUsuarioEnListadoAmigos(Usuario As String) As String
Dim Respuesta, Estadoamigo, Contador As Integer
Dim CadenaTemp, Sexo, UsuarioAlias, Comando, MensajeFinal, EstadoAmigoTexto As String
Dim IDVentana As Long
Dim Bandera, Usuarioexiste As Boolean
Dim TiempoInicial As Date
 
 ' **************************************************************
 ' Verificar si el usuario emisor existe en el listado de Amigos
 ' **************************************************************
 Bandera = False
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(Usuario)) Then
   Bandera = True
   Exit For
  End If
 Next
 
 If Bandera = False Then
  Respuesta = SolicitarAgregarAmigo(Usuario)
  UsuarioAlias = Trim(Variables.RespuestaAmigoACrearAlias)
  Estadoamigo = Variables.RespuestaAmigoACrearEstado
  EstadoAmigoTexto = Trim(Variables.RespuestaAmigoACrearEstadoTexto)
  Sexo = Trim(Variables.RespuestaAmigoACrearSexo)
  Usuarioexiste = True
  ' **************************************************************
  ' Como no existe pregunta si lo quiere agregar...
  
  Dim Resp As Long
  
  ' Si ya le pregunto y cancelo que no le pregunte... No Pregunta mas...
  If Varios.VolverAPreguntar(Trim(Usuario), "Buscar") = False Then
    Resp = Varios.MostrarMSGBox(MensajeRecurso(449) & Trim(Usuario) & MensajeRecurso(450), vbYesNo, "vbQuestion", Configuracion.TituloVentanas) ', True)
   Else
    VerificarUsuarioEnListadoAmigos = "NO"
    Exit Function ' Sale... No quiere agregarlo
  End If
  
  If Resp = vbNo Then
    ' Pregunta si quiere que se pregunte nuevamente...
    Resp = Varios.MostrarMSGBox(MensajeRecurso(467) & Trim(Usuario) & MensajeRecurso(468), vbYesNo, "vbQuestion", Configuracion.TituloVentanas) ', True)
    If Resp = vbNo Then
     Varios.VolverAPreguntar Trim(Usuario), "Agregar"
    End If
    VerificarUsuarioEnListadoAmigos = "NO"
    Exit Function ' Sale... No quiere agregarlo
  End If
  
  ' **************************************************************
  ' Verifica la Respuesta
  ' **************************************************************
  Select Case Respuesta
   Case -1
    ' No se Consiguio Respuesta del Servidor, ¿Desea Agregar al Amigo como 'Usuario Inexistente'?
    Respuesta = Varios.MostrarMSGBox(MensajeRecurso(136) & "[" & Trim(Variables.RespuestaAmigoACrearAlias) & "]" & MensajeRecurso(453), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
    If Respuesta = vbNo Then
     Exit Function
    End If
    Usuarioexiste = False
   Case 0
    ' El Amigo [ % ] no existe... ¿Desea Agregar al Amigo como 'Usuario Inexistente'?
    Respuesta = Varios.MostrarMSGBox(MensajeRecurso(134) & Trim(Variables.RespuestaAmigoACrearAlias) & MensajeRecurso(137), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
    If Respuesta = vbNo Then
     Exit Function
    End If
    Usuarioexiste = False
   Case 1
    ' Ok no hace nada y Agrega el Amigo, en estado desconectado
  End Select
  
  ' **************************************************************
  ' Crea el Nuevo Amigo
  ' **************************************************************
  Varios.CrearNuevoAmigo Usuarioexiste, "", CStr(UsuarioAlias), CInt(Estadoamigo), CStr(EstadoAmigoTexto), CStr(UCase(Sexo))
  'Variables.CantidadGrupoAmigo = Variables.CantidadGrupoAmigo + 1
  'ReDim Preserve Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
  'With Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
  ' .Existe = UsuarioExiste
  ' .NombreDelGrupo = ""
  ' .NombreDelAmigo = ""
  ' .IDNombreDelAmigo = UsuarioAlias
  ' .Sexo = UCase(Sexo)
  ' .EstadoDelAmigoEstado = EstadoAmigo
  ' .EstadoDelAmigoTexto = EstadoAmigoTexto
  'End With
  ' Espera 1 Segundo antes de hacer la grabacion...
  'TiempoInicial = Time
  'Do Until DateDiff("s", TiempoInicial, Time)
  ' DoEvents
  'Loop
  ' Grabar La Incorporacion...
  'Varios.GrabarListadoAmigos
 End If
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 
End Function

Public Function BuscarDireccionEmail(UsuarioID As String) As String
Dim Contador As Integer

 ' **************************************************************
 ' Busca el Usuario
 ' **************************************************************
 For Contador = 1 To CantidadGrupoAmigo
  If Trim(UCase(UsuarioID)) = Trim(UCase(GrupoAmigo(Contador).IDNombreDelAmigo)) Then
   ' Devuelve la direccion de Email
   BuscarDireccionEmail = GrupoAmigo(Contador).DireccionEMail
   Exit Function
  End If
 Next

 ' **************************************************************
 ' Si no encontro nada devuelve un Nulo...
 ' **************************************************************
 BuscarDireccionEmail = ""
 
End Function

Public Sub CentrarForm(Formulario As Form, Optional X As Boolean, Optional Y As Boolean)
Dim WidtH, HeighT As Long

   WidtH = Formulario.WidtH
   HeighT = Formulario.HeighT  ' Set height of form.
   ' Posicion X en Centro
   If X Then
    Formulario.Left = (Screen.WidtH - WidtH) / 2   ' Center form horizontally.
    Inicializar.PosicionCliente "Grabar", Formulario.Left, Formulario.Top
   End If
   ' Posicionar Y en centro
   If Y Then
    Formulario.Top = (Screen.HeighT - HeighT) / 2   ' Center form vertically.
    Inicializar.PosicionCliente "Grabar", Formulario.Left, Formulario.Top
   End If
   
End Sub
Public Sub EncajonarFormulario(Formulario As Form)
Dim Pantalla As POINTAPI
 
 Pantalla.X = Screen.WidtH
 Pantalla.Y = Screen.HeighT
 
  Formulario.Left = Pantalla.X - Formulario.WidtH
  Formulario.Top = Pantalla.Y - Formulario.HeighT
  Inicializar.PosicionCliente "Grabar", Formulario.Left, Formulario.Top
 
End Sub
Public Sub DefinirPosicionDeCliente(Optional Mostrar As Boolean)
Dim Posicion As POINTAPI
Dim PosicionTMP As POINTAPI

 ' **************************************************************
 ' Define la Posicion del Form
 ' **************************************************************
 Posicion = Inicializar.PosicionCliente("Leer")
 
 ' **************************************************************
 ' Si la posicion es afuera de la Pantalla lo
 ' Manda al Centro
 ' **************************************************************
 PosicionTMP.X = Screen.WidtH - Cliente.WidtH
 PosicionTMP.Y = Screen.HeighT - Cliente.HeighT
 
 ' **************************************************************
 ' Lo Posicion horizontalmente
 ' **************************************************************
 If Posicion.X > PosicionTMP.X Or Posicion.X < 0 Then
  CentrarForm Cliente, True
  Else
   'If Posicion.X < -3000 Then
   Cliente.Left = Posicion.X
   'End If
 End If
 
 ' **************************************************************
 ' Lo Posicion Verticalmente
 ' **************************************************************
 If Posicion.Y > PosicionTMP.Y Or Posicion.Y < 0 Then
  CentrarForm Cliente, , True
  Else
   'If Posicion.Y <> -10000 Then
    Cliente.Top = Posicion.Y
   'End If
 End If
 
 ' **************************************************************
 ' Muestra el Form...
 ' **************************************************************
 If Mostrar Then
  Cliente.Show
 End If
 
End Sub
Public Function BuscarElGrupoDelAmigo(UsuarioID) As String
Dim Contador As Integer

 For Contador = 1 To Variables.CantidadGrupoAmigo
  ' Busca el Usuario
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(UsuarioID)) Then
   BuscarElGrupoDelAmigo = Trim(GrupoAmigo(Contador).NombreDelGrupo)
   Exit Function
  End If
 Next
    
End Function
Public Function CambiarUsuarioDeGrupo(UsuarioID As String, Grupo As String, MandarRespuesta As Boolean) As Integer
Dim Contador, Respuesta As Integer

 ' **************************************************************
 ' Graba el Estado de los Nodos Actuales...
 ' **************************************************************
 CargarEstadoDeNodos
 
 ' **************************************************************
 ' Busca el usuario y lo cambia de grupo...
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  ' Busca el Usuario
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(UsuarioID)) Then
   ' Fuera % Fuera de Grupo...
   If UCase(Trim(Grupo)) = UCase(Trim(MensajeRecurso(148))) Or UCase(Trim(Grupo)) = UCase(Trim(MensajeRecurso(132))) Then
     GrupoAmigo(Contador).NombreDelGrupo = ""
    Else
     GrupoAmigo(Contador).NombreDelGrupo = Grupo
   End If
   If MandarRespuesta = True Then
    ' El Cambio del Amigo [ % ] al Grupo [ % ] fue Exitoso...
    Respuesta = MostrarMSGBox(MensajeRecurso(319) & Trim(UsuarioID) & MensajeRecurso(142) & Trim(Grupo) & MensajeRecurso(321), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
   End If
   ' Envia la Informacion del Nuevo Listado de Amigos
   Varios.CargarAmigos
   ' Envia la Informacion del Nuevo Listado de Amigos
   Varios.GrabarListadoAmigos
   CambiarUsuarioDeGrupo = 1 ' OK
   Exit Function
  End If
 Next
 
 CambiarUsuarioDeGrupo = 0 ' Hubo un Error
 ' Hubo un problema el Intentar Cambiar el Amigo [ % ] al Grupo [
 Respuesta = MostrarMSGBox(MensajeRecurso(322) & Trim(UsuarioID) & MensajeRecurso(142) & Trim(Grupo) & "...", vbOKOnly, "vbCritical", Configuracion.TituloVentanas)

End Function
Public Function NuevoID() As String

 ' **************************************************************
 ' Genera un nuevo ID de Mensaje
 ' **************************************************************
 IDMensaje = IDMensaje + 1
 
 ' **************************************************************
 ' Verifica que no se pase
 ' **************************************************************
 If IDMensaje = 100 Then IDMensaje = 1
 
 ' **************************************************************
 ' Devuelve el ID
 ' **************************************************************
 NuevoID = CompletarCadena(CStr(IDMensaje), 2, "I", "0")
 
End Function

Public Sub EsperarSegundos(CantidadDeSegundos As Long)
'Dim Contador As Long
Dim Inicio As Date

 ' **************************************************************
 ' Espera xx sgundos
 ' **************************************************************
 Inicio = Time
 Do Until DateDiff("s", Inicio, Time) >= CantidadDeSegundos
  DoEvents
 Loop
 
End Sub

Public Sub HabilitarDeshabilitarBloqueo(Usuario As String, Bloqueado As Boolean)
Dim Contador, Respuesta As Integer

 ' **************************************************************
 ' Cambia en Los Formularios de Mensaje el Estado de los Amigos
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
   ' Busca el Formulario de Mensaje
  If Forms(Contador).FormularioNombre = "Mensajes" Then
    ' Verifica cuantos Usuarios tiene
    If Forms(Contador).CantidadDeAmigosEnChat = 1 Then
      ' Verificar que el Chat sea con el Usuario
      Respuesta = Forms(Contador).TratarAmigosEnChat("Buscar", CStr(Usuario))
      If Respuesta = 1 Then
       Forms(Contador).ImagenBloqueado.Visible = Bloqueado
      End If
     Else
      ' Sino comom es MultiChat lo esconde
      Forms(Contador).ImagenBloqueado.Visible = False
    End If
  End If
 Next
  
End Sub
Public Sub EfectoBoton(ShapeNombre As Shape, Optional Color As Long)
Dim Hora As Date
Dim ColorActual, Contador As Long

 ' **************************************************************
 ' Si el Color es Nulo, entonces es Gris Oscuro
 ' **************************************************************
 If IsNull(Color) Or Color = 0 Then
  Color = &HC0C0C0
 End If
 
 ' **************************************************************
 ' Colores actuales
 ' **************************************************************
 ColorActual = ShapeNombre.BackColor
 Hora = Time
 
 ' **************************************************************
 ' Cambia el Color
 ' **************************************************************
 ShapeNombre.BackColor = Color
 
 ' **************************************************************
 ' Espera un Segundo
 ' **************************************************************
 'Do
 ' DoEvents
 ' If DateDiff("s", Hora, Time) > 1 Then
 '  ShapeNombre.BackColor = ColorActual
 '  Exit Do
 ' End If
 'Loop
 
 ' **************************************************************
 ' Devuelve el Color
 ' **************************************************************
 ShapeNombre.BackColor = ColorActual
 
End Sub
Public Sub CargarFuentes()
Dim Contador As Integer
Dim Cantidad As Integer

 ' **************************************************************
 ' Carga las Fuentes del Servicio
 ' **************************************************************
 ReDim Variables.NombreDeFuentes(Screen.FontCount)
 Cantidad = 0
 For Contador = 0 To Screen.FontCount - 1
  DoEvents ' Permite que siga la ejecucion de otros procesos
  If Trim(Screen.Fonts(Contador)) <> "" Then
   Cantidad = Cantidad + 1
   Variables.NombreDeFuentes(Cantidad) = Trim(Screen.Fonts(Contador))
  End If
 Next Contador
 
 Variables.CantidadDeFuentes = Cantidad
 
End Sub
Public Sub ProcesarUsuariosBloqueados(Metodo As String, Nombre As String)
Dim Contador, Numero As Integer
Dim Existe As Boolean

 Select Case Metodo
    Case "Agregar"
     If Variables.UsuarioBloqueadoCantidad > 0 Then
      ' **************************************************************
      ' Primero Verifica que no Exista...
      ' **************************************************************
      Existe = False
      For Contador = 1 To Variables.UsuarioBloqueadoCantidad
       If UCase(Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)) = UCase(Trim(Nombre)) Then
        Existe = True
        Exit Sub
       End If
      Next
     End If
     ' **************************************************************
      
     ' **************************************************************
     ' Procesa el Nuevo Amigo
     ' **************************************************************
     Variables.UsuarioBloqueadoCantidad = Variables.UsuarioBloqueadoCantidad + 1
     ReDim Preserve Variables.UsuarioBloqueadoNombres(Variables.UsuarioBloqueadoCantidad)
     Variables.UsuarioBloqueadoNombres(Variables.UsuarioBloqueadoCantidad).NombreDelUsuario = Trim(Nombre)
     
     ' **************************************************************
     ' Cambiar en Todas las Ventans
     ' **************************************************************
     HabilitarDeshabilitarBloqueo Trim(Nombre), True
     
     ' **************************************************************
     ' Graba los Cambios
     ' **************************************************************
     Inicializar.GrabarBloqueUsuarios
    
    Case "Sacar"
     ' **************************************************************
     ' Primero Verifica si Existe
     ' **************************************************************
     Numero = 0
     For Contador = 1 To Variables.UsuarioBloqueadoCantidad
      If UCase(Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)) = UCase(Trim(Nombre)) Then
       Numero = Contador
       Exit For
      End If
     Next
     ' **************************************************************
     ' El Amigo que se solicita sacar no existe..
     If Contador = 0 Then Exit Sub
     
      ' **************************************************************
      ' Cambiar en Todas las Ventans
      ' **************************************************************
      HabilitarDeshabilitarBloqueo Trim(Nombre), False
     
     ' **************************************************************
     ' Procesa el Nuevo Amigo
     ' **************************************************************
     ' Si es el Unico lo Borra
     If Variables.UsuarioBloqueadoCantidad = 1 Then
      Variables.UsuarioBloqueadoCantidad = 0
      ' **************************************************************
      ' Graba los Cambios
      ' **************************************************************
      Inicializar.GrabarBloqueUsuarios
      Exit Sub
     End If
     
     ' Si es el Ultimo Resta 1
     If Variables.UsuarioBloqueadoCantidad = Contador Then
      Variables.UsuarioBloqueadoCantidad = Variables.UsuarioBloqueadoCantidad - 1
      ' **************************************************************
      ' Graba los Cambios
      ' **************************************************************
      Inicializar.GrabarBloqueUsuarios
      Exit Sub
     End If
     
     ' Sino copia el Ultimo a la Posicion del Actual...
     'If Variables.UsuarioBloqueadoCantidad = Contador Then
     Variables.UsuarioBloqueadoNombres(Contador) = Variables.UsuarioBloqueadoNombres(Variables.UsuarioBloqueadoCantidad)
     Variables.UsuarioBloqueadoCantidad = Variables.UsuarioBloqueadoCantidad - 1
     ' **************************************************************
     ' Graba los Cambios
     ' **************************************************************
     Inicializar.GrabarBloqueUsuarios
     ' Exit Sub
     'End If
        
 End Select
 
End Sub
Public Sub CerrarVentanasDeMenus()
Dim Contador As Integer

 ' **************************************************************
 ' Cierra todas las Ventanas de Menu en el Sistema...
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
  If Forms(Contador).FormularioNombre = "VentanaMenu" Then
   If Forms(Contador).Visible = True Then
    Forms(Contador).Hide
   End If
  End If
 Next
 
End Sub
Function CrearVentanaMensajeOffLine(AliasUsuario As String, Para As String)
Dim MensajeriaOffLine As New MensajeOffLine
   
 ' **************************************************************
 ' Crea la Ventana de Mensaje...
 ' **************************************************************
 Load MensajeriaOffLine
 MensajeriaOffLine.AliasUsuario = AliasUsuario
 MensajeriaOffLine.MensajeOfflinePara = Para
 MensajeriaOffLine.MensajeOfflinePara.Enabled = False
 MensajeriaOffLine.Show
 
End Function

Function CrearVentanaMensaje(Amigo As String) As Integer
Dim Mensajeria As New Mensajes
Dim Handle As Long

 ' **************************************************************
 ' Crea la Ventana de Mensaje...
 ' **************************************************************
 Load Mensajeria
 Mensajeria.TratarAmigosEnChat "Agregar", Amigo
 Mensajeria.MostrarVentana
 Handle = Mensajeria.hwnd '''
 'CrearVentanaMensaje = Forms.Count - 1
 CrearVentanaMensaje = BuscarVentanaHandle(Handle) '
 
End Function
Sub CambiarIconoTray(Icono As String)
Dim MensajeDelIconTray As String

 ' **************************************************************
 ' Define el Texto del ToolTip del IconTray
 ' **************************************************************
 ' Estado Disponible
 If CInt(Configuracion.EstadoDelUsuario) = 1 Then MensajeDelIconTray = Trim(Configuracion.IDAliasUsuario) & " - " & MensajeRecurso(180)
 ' Estado No Disponible
 If CInt(Configuracion.EstadoDelUsuario) = 2 Then MensajeDelIconTray = Trim(Configuracion.IDAliasUsuario) & " - " & MensajeRecurso(181)
 ' Estado Custom y/o Otros...
 If CInt(Configuracion.EstadoDelUsuario) = 3 Then MensajeDelIconTray = Trim(Configuracion.IDAliasUsuario) & " - " & Trim(Configuracion.EstadoActualTexto)
  
 ' **************************************************************
 ' Define el Icono a Mostrar en el Icon Tray
 ' **************************************************************
 Select Case UCase(Icono)
  ' **************************************************************
  ' Desconectado
  ' **************************************************************
  Case UCase("DesConectado")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectado.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensaje")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensaje.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash2")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash2.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash3")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash3.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash4")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash4.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash5")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash5.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash6")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectadoConMensajeFlash6.Picture, Trim(Configuracion.TituloVentanas)
  Case UCase("DesConectadoConMensajeFlash0")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayDesConectado.Picture, Trim(Configuracion.TituloVentanas)
  
  ' **************************************************************
  ' Conectando
  ' **************************************************************
  Case UCase("Conectando")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectando.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensaje")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensaje.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash2")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash2.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash3")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash3.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash4")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash4.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash5")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash5.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  Case UCase("ConectandoConMensajeFlash6")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectandoConMensajeFlash6.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  
  Case UCase("ConectandoConMensajeFlash0")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectando.Picture, Trim(Configuracion.TituloVentanas) & " - " & Trim(Configuracion.IDAliasUsuario) & "..."
  
  ' **************************************************************
  ' Conectado
  ' **************************************************************
  Case UCase("ConectadoSinMensaje")
   ' **************************************************************
   ' Segun el Estado de el Usuario Define la Imagen a Mostrar...
   ' **************************************************************
   ' Estado Conectado
   If Configuracion.EstadoDelUsuario = 1 Then Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoSinMensajeVerde.Picture, MensajeDelIconTray
   ' Estado No Disponible
   If Configuracion.EstadoDelUsuario = 2 Then Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoSinMensajeRojo.Picture, MensajeDelIconTray
   ' Estado Custom y/o Otros...
   If Configuracion.EstadoDelUsuario = 3 Then Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoSinMensajeAmarillo.Picture, MensajeDelIconTray
   ' **************************************************************
  Case UCase("ConectadoConMensaje")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensaje.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash2")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash2.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash3")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash3.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash4")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash4.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash5")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash5.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash6")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoConMensajeFlash6.Picture, MensajeDelIconTray
  Case UCase("ConectadoConMensajeFlash0")
   Cliente.SysIcon.CambiarIconoSystemTray Cliente.IconoTrayConectadoSinMensaje.Picture, MensajeDelIconTray
 End Select

End Sub
Sub AgregarMensajesPendientes(MensajeDe As String, Mensaje As String, Optional HoraYFecha As String)

 ' **************************************************************
 ' Agrega Mensajes Pendientes...
 ' **************************************************************
 Variables.CantidadDeMensajesPendientes = Variables.CantidadDeMensajesPendientes + 1
 ReDim Preserve Variables.MensajesPendientes(Variables.CantidadDeMensajesPendientes)
 With Variables.MensajesPendientes(Variables.CantidadDeMensajesPendientes)
  .MensajeDe = Trim(MensajeDe)
  .Mensaje = Trim(Mensaje)
  If Trim(HoraYFecha) = "" Then
    .HoraYFecha = Time & "_" & Date
   Else
    .HoraYFecha = HoraYFecha
  End If
 End With
 CambiarMensajesPendientes (Variables.CantidadDeMensajesPendientes)
  
 ' **************************************************************
 ' Graba los Mensajes Pendientes
 ' **************************************************************
 GrabarMensajesPendientes
 
End Sub
Sub CargarMensajesPendientes()
Dim Contador, CantidadDeMensajes As Integer
Dim Archivo As String
Dim MensajePendiente As MiMensajesPendientesFormatoDefinido
Dim UsuarioDir As String

 ' **********************************************************************
 ' Captura el Error...
 ' **********************************************************************
 On Error GoTo SalirCargarMensajesPendientes
 
 ' **********************************************************************
 ' Carga los Mensajes Pendientes...
 ' **********************************************************************
 UsuarioDir = "Software\Eim\Mensajes\" & Trim(UCase(Configuracion.IDAliasUsuario))
 ' Carga la Cantidad de mensajes que hay pendientes...
 CantidadDeMensajes = LeerRegistry(HKEY_LOCAL_MACHINE, UsuarioDir, "CantidadDeMensajes")
 ' Si la cantidad es nulo o cero, sale sin hacer nada
 If CantidadDeMensajes = 0 Or IsNull(CantidadDeMensajes) Then
  Varios.CambiarMensajesPendientes (0)
 End If
 ' Leer los Mensajes Pendientes
 ReDim Preserve Variables.MensajesPendientes(CantidadDeMensajes)
 For Contador = 1 To CantidadDeMensajes
   With MensajesPendientes(Contador)
    .HoraYFecha = LeerRegistry(HKEY_LOCAL_MACHINE, UsuarioDir, "HoraYFecha" & Contador)
    .Mensaje = LeerRegistry(HKEY_LOCAL_MACHINE, UsuarioDir, "Mensaje" & Contador)
    .MensajeDe = LeerRegistry(HKEY_LOCAL_MACHINE, UsuarioDir, "MensajeDe" & Contador)
   End With
 Next
 ' Define la Cantidad Final de Mensajes
 Varios.CambiarMensajesPendientes (CantidadDeMensajes)
 
 
SalirCargarMensajesPendientes:

End Sub
Sub GrabarMensajesPendientes()
Dim Contador, CantidadDeMensajes As Integer
Dim Archivo, CantidadTMP As String
Dim MensajePendiente As MiMensajesPendientesFormatoDefinido
Dim UsuarioDir As String

 ' **********************************************************************
 ' Captura el Error...
 ' **********************************************************************
 On Error GoTo SalirGrabarMensajesPendientes
 
 ' **********************************************************************
 ' Define la Key donde se guardan los mensajes del usuario
 ' **********************************************************************
 UsuarioDir = "Software\Eim\Mensajes\" & Trim(UCase(Configuracion.IDAliasUsuario))
 
 ' **********************************************************************
 ' Borra los Viejos Mensajes
 ' **********************************************************************
 CantidadTMP = LeerRegistry(HKEY_LOCAL_MACHINE, UsuarioDir, "CantidadDeMensajes")
 ' Si hay mas de 0 mensajes los borra
 If IsNumeric(CantidadTMP) Then
   CantidadDeMensajes = CantidadTMP
   If CantidadDeMensajes <> 0 Then
    For Contador = 1 To CantidadDeMensajes
      With MensajesPendientes(Contador)
       GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "HoraYFecha" & Contador, ""
       GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "Mensaje" & Contador, ""
       GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "MensajeDe" & Contador, ""
      End With
    Next
   End If
  Else
   CantidadDeMensajes = 0
 End If
 
 
 ' **********************************************************************
 ' Graba el Mensaje
 ' **********************************************************************
 ' Graba o Crea la key del Usuario
 GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "CantidadDeMensajes", CStr(CantidadDeMensajesPendientes)
 ' **********************************************************************
 ' Graba la cantidad, si es 0 no graba ningun mensaje
 ' **********************************************************************
 If Variables.CantidadDeMensajesPendientes = 0 Then Exit Sub
 ' Graba Los Mensajes
 For Contador = 1 To CantidadDeMensajesPendientes
   With MensajesPendientes(Contador)
    GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "HoraYFecha" & Contador, .HoraYFecha
    GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "Mensaje" & Contador, .Mensaje
    GrabarRegistry HKEY_LOCAL_MACHINE, UsuarioDir, "MensajeDe" & Contador, .MensajeDe
   End With
 Next
 
SalirGrabarMensajesPendientes:

End Sub
Sub AgruparMensajesPendientes()
Dim UsuarioEmisor As String
Dim CantidadDeEmisores As Integer
Dim Contador, Contador2 As Integer
Dim Bandera As Boolean

 ' **************************************************************
 ' Si no hay mensajes pendientes sale sin hacer nada...
 ' **************************************************************
 If CantidadDeMensajesPendientes = 0 Then Exit Sub
 
 ' **************************************************************
 ' Si hay uno solo lo pasa directo
 ' **************************************************************
 If CantidadDeMensajesPendientes = 1 Then
  ReDim Variables.MensajesPendientesAgrupados(CantidadDeMensajesPendientes)
  Variables.CantidadDeMensajesPendientesAgrupados = 1
  Variables.MensajesPendientesAgrupados(1).MensajeDe = Variables.MensajesPendientes(1).MensajeDe
  Exit Sub
 End If
 
 ' **************************************************************
 ' Primero genera un Listado de los Emisores
 ' **************************************************************
 CantidadDeEmisores = 1
 ' Lo redimensiona a un Maximo que nunca puede ser pasado, ya que siempre
 ' la cantidad de emisores tiene que ser menor o igual a la cantidad de
 ' mensajes pendientes....
 ReDim Variables.MensajesPendientesAgrupados(CantidadDeMensajesPendientes)
 For Contador = 1 To CantidadDeMensajesPendientes
  UsuarioEmisor = MensajesPendientes(Contador).MensajeDe
  Bandera = False
  For Contador2 = 1 To CantidadDeMensajesPendientes
   ' Si encuentra alguno igual ya es suficiente para no agregarlo
   If Trim(UCase(Variables.MensajesPendientesAgrupados(Contador2).MensajeDe)) = Trim(UCase(UsuarioEmisor)) Then
    Bandera = True
   End If
  Next
  ' Agrega el Emisor
  If Bandera = False Then
   Variables.MensajesPendientesAgrupados(CantidadDeEmisores).MensajeDe = UsuarioEmisor
   CantidadDeEmisores = CantidadDeEmisores + 1
  End If
 Next
 
 ' **************************************************************
 ' Si cantidad de Emisores es 0 Sale
 ' **************************************************************
 CantidadDeEmisores = CantidadDeEmisores - 1
 If CantidadDeEmisores = 0 Then Exit Sub
 
 ' **************************************************************
 ' Define la Cantidad de Emisores
 ' **************************************************************
 Variables.CantidadDeMensajesPendientesAgrupados = CantidadDeEmisores
 
 ' **************************************************************
 ' Carga la Cantidad de Mensajes de Cada Emisor
 ' **************************************************************
 For Contador = 1 To Variables.CantidadDeMensajesPendientesAgrupados
  MensajesPendientesAgrupados(Contador).Cantidad = 0
  For Contador2 = 1 To CantidadDeMensajesPendientes
   If Trim(UCase(MensajesPendientesAgrupados(Contador).MensajeDe)) = Trim(UCase(MensajesPendientes(Contador2).MensajeDe)) Then
    MensajesPendientesAgrupados(Contador).Cantidad = MensajesPendientesAgrupados(Contador).Cantidad + 1
   End If
  Next
 Next
 
End Sub
Sub CambiarMensajesPendientes(Cantidad As Integer)

 ' **************************************************************
 ' Cambia la Cantidad de Mensajes Pendientes
 ' **************************************************************
 Variables.CantidadDeMensajesPendientes = Cantidad
 
 ' **************************************************************
 ' Verifica que existe el Nodo que debe Cambiarse
 ' **************************************************************
 If Cliente.ListadoDeAmigos.Nodes.Count = 0 Then
  Exit Sub
 End If
 If Mid$(Cliente.ListadoDeAmigos.Nodes(1).Text, 1, 1) <> "*" Then
  Exit Sub
 End If
 
 
 Select Case Cantidad
  Case 0
   Cliente.TimerMensaje.Enabled = False
   Cliente.IndiceTimerMensaje = 0
   ' * Usted Tiene (0) Mensajes Pendientes...
   Cliente.ListadoDeAmigos.Nodes(1).Text = MensajeRecurso(324)
   Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash3"
   Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash3"
   Cliente.ListadoDeAmigos.Nodes(1).Bold = False
   Cliente.ListadoDeAmigos.Nodes(1).ForeColor = vbBlack
   Cliente.ListadoDeAmigos.Nodes(1).Expanded = True
   ' **************************************************************
   ' Si esta desconectado o Conectando el Icon Tray
   ' no cambia su Estado...
   ' **************************************************************
   If Configuracion.Logueado <> 0 And Configuracion.Logueado <> 1 Then
    Varios.CambiarIconoTray "ConectadoSinMensaje"
   End If
  Case 1
   If Cliente.TimerMensaje.Enabled = False Then
    Audio.EjecutarSonido "004"
    Cliente.TimerMensaje.Enabled = True
    Cliente.IndiceTimerMensaje = 0
   End If
   ' * Usted Tiene (1) Mensaje Pendiente...
   Cliente.ListadoDeAmigos.Nodes(1).Text = MensajeRecurso(325)
   Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash"
   Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash"
   Cliente.ListadoDeAmigos.Nodes(1).Bold = False
   Cliente.ListadoDeAmigos.Nodes(1).ForeColor = vbBlue
   Cliente.ListadoDeAmigos.Nodes(1).Expanded = True
   ' **************************************************************
   ' Si esta desconectado o Conectando el Icon Tray
   ' no cambia su Estado...
   ' **************************************************************
   If Configuracion.Logueado <> 0 And Configuracion.Logueado <> 1 Then
    Varios.CambiarIconoTray "ConectadoConMensaje"
   End If
  Case Else
   If Cliente.TimerMensaje.Enabled = False Then
    Audio.EjecutarSonido "004"
    Cliente.TimerMensaje.Enabled = True
    Cliente.IndiceTimerMensaje = 0
   End If
   ' * Usted Tiene ( % ) Mensaje Pendiente...
   Cliente.ListadoDeAmigos.Nodes(1).Text = MensajeRecurso(326) & Cantidad & MensajeRecurso(327)
   Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash"
   Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash"
   Cliente.ListadoDeAmigos.Nodes(1).Bold = False
   Cliente.ListadoDeAmigos.Nodes(1).ForeColor = vbBlue
   Cliente.ListadoDeAmigos.Nodes(1).Expanded = True
   ' **************************************************************
   ' Si esta desconectado o Conectando el Icon Tray
   ' no cambia su Estado...
   ' **************************************************************
   If Configuracion.Logueado <> 0 And Configuracion.Logueado <> 1 Then
    Varios.CambiarIconoTray "ConectadoConMensaje"
   End If
 End Select
 
End Sub
Sub SalirDelSistema()
Dim Respuesta As Integer

  ' ¿Está Seguro que Desea Salir del Sistema...?
  Respuesta = MostrarMSGBox(MensajeRecurso(179), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
  
  ' Cancela la salida del Sistema
  If Respuesta = vbNo Then
   Exit Sub
  End If
  
  ' Cierra Todo
  Cliente.SysIcon.HideIcon
  Cliente.TCPSocket.Close
  End ' Cierra Todo
  
End Sub
Sub EliminarGrupo(GrupoAEliminar As String)
Dim Grupo As String
Dim NuevaCantidad, Respuesta, Contador As Integer
Dim GrupoAmigoTMP() As MiGrupoAmigo

 ' **************************************************************
 ' Confirma que se quiera borrar del Grupo
 ' **************************************************************
 Grupo = GrupoAEliminar
 ' ¿Está Seguro que Desea Borrar el Grupo [
 Respuesta = MostrarMSGBox(MensajeRecurso(329) & Grupo & "]?", vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
  
 ' El Borrado del Grupo
 If Respuesta = vbNo Then
  Exit Sub
 End If

 ' Borra el Grupo de todos los Amigos
 NuevaCantidad = 0
 ReDim GrupoAmigoTMP(Variables.CantidadGrupoAmigo)
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).NombreDelGrupo)) = UCase(Grupo) Then
   GrupoAmigo(Contador).NombreDelGrupo = ""
  End If
  ' Pasa a una variable temporal el Listado Actual sin los borrados
  If Trim(UCase(Trim(GrupoAmigo(Contador).NombreDelGrupo))) <> "" Or Trim(UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo))) <> "" Then
   NuevaCantidad = NuevaCantidad + 1
   GrupoAmigoTMP(NuevaCantidad) = GrupoAmigo(Contador)
  End If
 Next
 
 ' Deja solo los Usuarios y Grupos Validos (De acuerdo a lo ya Procesado...)
 For Contador = 1 To NuevaCantidad
  GrupoAmigo(Contador) = GrupoAmigoTMP(Contador)
 Next
 Variables.CantidadGrupoAmigo = NuevaCantidad
 
 ' **************************************************************
 ' Graba el Estado de los Nodos Actuales...
 ' **************************************************************
 CargarEstadoDeNodos
 
 ' **************************************************************
 ' Envia la Informacion del Nuevo Listado de Amigos
 ' **************************************************************
 Varios.CargarAmigos
 
 ' **************************************************************
 ' Envia la Informacion del Nuevo Listado de Amigos
 ' **************************************************************
 Varios.GrabarListadoAmigos
 
End Sub
Sub RenombrarGrupo(GrupoARenombrar As String)
Dim Grupo, Respuesta As String
Dim Contador, Resp As Integer
Dim GrupoAmigoTMP() As MiGrupoAmigo

 ' **************************************************************
 ' Define el Nuevo nombre del Grupo
 ' **************************************************************
 Dim Bandera As Boolean
 Do
  Grupo = GrupoARenombrar
  Bandera = False
  ' Por Favor Ingrese el Nuevo Nombre del Grupo '
  Respuesta = MostrarInputBox(MensajeRecurso(330) & GrupoARenombrar & "]...", 20, Configuracion.TituloVentanas)
  If Respuesta = "" Then Exit Sub
  For Contador = 1 To Variables.CantidadGrupoAmigo
   If UCase(Trim(GrupoAmigo(Contador).NombreDelGrupo)) = UCase(Respuesta) Then
    Bandera = True
   End If
  Next
   If Bandera = True Then
     ' El Grupo ' % ' ya existe... ¿Desea Cambiar el Nombre del Grupo a otro?
     Resp = MostrarMSGBox(MensajeRecurso(331) & Trim(Respuesta) & MensajeRecurso(332), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
     If Resp = vbNo Then Exit Sub ' Se equivoco pero no quiere seguir...
    Else
     If Trim(Respuesta) <> "" Then
       Exit Do  ' Sigue para el Cambio de Nombre
      Else
       Exit Sub ' Como no hubo una respuesta valida sale...
     End If
   End If
   ' Aca vuelve a preguntar el Nombre
 Loop
 
 ' Borra el Grupo de todos los Amigos
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).NombreDelGrupo)) = UCase(Grupo) Then
   GrupoAmigo(Contador).NombreDelGrupo = Respuesta
  End If
 Next
 
 ' **************************************************************
 ' Graba el Estado de los Nodos Actuales...
 ' **************************************************************
 CargarEstadoDeNodos (Respuesta)

 ' Envia la Informacion del Nuevo Listado de Amigos
 Varios.CargarAmigos
 
 ' Envia la Informacion del Nuevo Listado de Amigos
 Varios.GrabarListadoAmigos

End Sub

Sub EliminarAmigo(AmigoAEliminar As String)
Dim Respuesta, Contador As Integer
Dim Amigo As String

 ' **************************************************************
 ' Confirma que se quiera borrar el Amigo
 ' **************************************************************
 Amigo = Trim(AmigoAEliminar)
 ' ¿Está Seguro que Desea Borrar a [ % ] de su Listado de Amigos?
 Respuesta = MostrarMSGBox(MensajeRecurso(333) & Amigo & MensajeRecurso(334), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
  
 ' El Borrado de Amigos
 If Respuesta = vbNo Then
  Exit Sub
 End If

 ' Borra el Amigo
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Amigo) Then
   ' Si es el Ultimo Item Simplemente pone la Cantidad de Items en
   ' CantidadGrupoAmigo -1
   If Contador = Variables.CantidadGrupoAmigo Then
     Variables.CantidadGrupoAmigo = Variables.CantidadGrupoAmigo - 1
     Exit For
    Else
     ' Sino copia el Ultimo a esta posicion, y resta uno a la cantidad
     ' de CantidadGrupoAmigo
     GrupoAmigo(Contador) = GrupoAmigo(Variables.CantidadGrupoAmigo)
     Variables.CantidadGrupoAmigo = Variables.CantidadGrupoAmigo - 1
    Exit For
   End If
  End If
 Next
 
 ' **************************************************************
 ' Graba el Estado de los Nodos Actuales...
 ' **************************************************************
 CargarEstadoDeNodos

 ' Envia la Informacion del Nuevo Listado de Amigos
 Varios.CargarAmigos
 
 ' Envia la Informacion del Nuevo Listado de Amigos
 Varios.GrabarListadoAmigos
 
End Sub
Sub CrearGrupo()
Dim GrupoACrear As String
Dim Contador As Integer

 ' **************************************************************
 ' Solicita Se Ingrese el Nombre del Grupo
 ' **************************************************************
 ' Por Favor Ingrese el Nombre del Grupo Que Desea Crear:
 GrupoACrear = Varios.MostrarInputBox(MensajeRecurso(335), 20, Configuracion.TituloVentanas)
 
 ' **************************************************************
 ' Verifica el GrupoACrear
 ' **************************************************************
 If Trim(GrupoACrear) = "" Then Exit Sub
 

 ' **************************************************************
 ' Verificar que el Grupo en Cuestion no existe...
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(Variables.GrupoAmigo(Contador).NombreDelGrupo)) = UCase(GrupoACrear) Then
   ' El Grupo [ % ] ya Existe...
   MostrarMSGBox MensajeRecurso(331) & Trim(GrupoACrear) & MensajeRecurso(337), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Exit Sub
  End If
 Next
 
 ' **************************************************************
 ' Crea el Nuevo Grupo
 ' **************************************************************
 Variables.CantidadGrupoAmigo = Variables.CantidadGrupoAmigo + 1
 ReDim Preserve Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
 With Variables.GrupoAmigo(Variables.CantidadGrupoAmigo)
  .Existe = True
  .NombreDelGrupo = GrupoACrear
  .NombreDelAmigo = ""
  .IDNombreDelAmigo = ""
 End With
 
 ' **************************************************************
 ' Grabar la Eliminacio
 ' **************************************************************
 Varios.GrabarListadoAmigos

End Sub
Sub AgregarAmigo(NombreDelGrupo As String)

 Load AgregarBuscarAmigo
 AgregarBuscarAmigo.NombreDelGrupo = NombreDelGrupo
 If Trim(NombreDelGrupo) <> "" Then
  AgregarBuscarAmigo.Grupo = Trim(NombreDelGrupo)
 End If
 AgregarBuscarAmigo.Show 'vbModal
 
End Sub
Sub CargarEstadoDeNodos(Optional Nombre As String)
Dim Contador As Integer

 ' **************************************************************
 ' Si viene 1 Nombre agrege el Mismo al Listado... Esto es usado
 ' para renombrar un Grupo...
 ' **************************************************************
 If Trim(Nombre) <> "" Then
  GrupoEstadoCantidad = GrupoEstadoCantidad + 1
  ReDim Preserve GrupoEstadoNombre(GrupoEstadoCantidad)
  GrupoEstadoNombre(GrupoEstadoCantidad) = UCase(Nombre)
  ' Graba en registry el Estado de los Grupos...
  Inicializar.GrabarEstadoDeGrupos
  Exit Sub
 End If
 
 ' **************************************************************
 ' Carga los Grupos actuales con el estado (Expandido o no)
 ' **************************************************************
 ' Solo si es mayor que 2 ya que si esta en dos es por que el
 ' listado tienen el mensaje de refrescando amigos...
 If Cliente.ListadoDeAmigos.Nodes.Count > 2 Then
  GrupoEstadoCantidad = 0
  For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
   If UCase(Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1)) = "G" And Trim(UCase(Cliente.ListadoDeAmigos.Nodes(Contador).key)) <> "USUARIO" Then
    If Cliente.ListadoDeAmigos.Nodes(Contador).Expanded = False Then
     GrupoEstadoCantidad = GrupoEstadoCantidad + 1
     ReDim Preserve GrupoEstadoNombre(GrupoEstadoCantidad)
     GrupoEstadoNombre(GrupoEstadoCantidad) = UCase(Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2))
    End If
   End If
  Next
 End If
 ' Graba en registry el Estado de los Grupos...
 Inicializar.GrabarEstadoDeGrupos
 
End Sub
Sub CambiaEstadoListadoDeAmigos(Estado As Integer)
Dim Nodo As Node
Dim UsuarioNombre As String

 ' **************************************************************
 ' Carga que nodos estan expandidos y cuales no...
 ' **************************************************************
 ' La Variable GrupoEstadoPrimeraVez se utiliza para determinar
 ' que ya se cargo de Registry el estado de los grupos, ya que
 ' se carga cuando el estado del usuario pasa a 3 (Logueado)
 If Variables.GrupoEstadoPrimeraVez Then
   Variables.GrupoEstadoPrimeraVez = False
  Else
   CargarEstadoDeNodos
 End If
 
 ' **************************************************************
 ' Define como se ve el Listado de Amigos cuando esta:
 ' Desconectado y Conectando...
 ' **************************************************************
 LimpiarListadoDeAmigos
 'Cliente.ListadoDeAmigos.Nodes.Clear
 UsuarioNombre = UCase(Mid$(Trim(Variables.Configuracion.IDAliasUsuario), 1, 1)) & Mid$(Trim(Variables.Configuracion.IDAliasUsuario), 2)
   
 ' **************************************************************
 ' Crea el Nodo de Mensajes Pendientes...
 ' **************************************************************
 ' * Usted Tiene (0) Mensajes Pendientes...
 Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "MensajesPendientes", MensajeRecurso(324), "MensajeFlash3", "MensajeFlash3")
 Nodo.Bold = False
 Nodo.ForeColor = vbBlack
 Nodo.BackColor = vbWhite
 Nodo.Expanded = True
 
 ' **************************************************************
 ' Crea el Nodo del usuario...
 ' **************************************************************
 Select Case Estado
  Case 0
   ' Desconectado...
   Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(228) & "...", "Desconectado", "Desconectado")
   ' Lo deja Abierto para poder ver los Mensajes Pendientes...
   ' Cliente.ListadoDeAmigos.Enabled = False
  Case 1
   ' Conectando...
   Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(229) & "...", "Conectando", "Conectando")
   ' Lo deja Abierto para poder ver los Mensajes Pendientes...
   ' Cliente.ListadoDeAmigos.Enabled = False
  Case 3
   ' Cargando Listado de [
   Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(338) & UsuarioNombre & "]...", "CargandoAmigos", "CargandoAmigos")
   Cliente.ListadoDeAmigos.Enabled = True
 End Select
 
 'Nodo.EnsureVisible
 Nodo.Bold = True
 Nodo.ForeColor = vbBlack
 Nodo.Expanded = True
 
End Sub
Sub RecargarListadoDeAmigos()
Dim UsuarioNombre As String
Dim Nodo As Node
   
 ' **************************************************************
 ' Recarga el Listado de Amigos, e Informa de Esta
 ' Situacion en ListadoDeAmigos...
 ' **************************************************************
 Varios.CambiaEstadoListadoDeAmigos (3)
 EnviarPaqueteTCP ("22")
    
End Sub
Sub GrabarListadoAmigos()
Dim Contador As Integer
Dim ListadoAEnviar As String

 ' **************************************************************
 ' Graba el Estado de los Nodos Actuales...
 ' **************************************************************
 CargarEstadoDeNodos

 ' **************************************************************
 ' Carga el Listado de Amigos Actual en caso de Que se hallan
 ' realizado cmabios...
 ' **************************************************************
 CargarAmigos
 
 ' **************************************************************
 ' Genera el Paquete a Enviar conteniendo el listado de
 ' Amigos
 ' **************************************************************
 ListadoAEnviar = ""
 For Contador = 1 To Variables.CantidadGrupoAmigo
  With GrupoAmigo(Contador)
  ListadoAEnviar = ListadoAEnviar & _
                    Trim(.IDNombreDelAmigo) & "@" & _
                    Trim(.NombreDelGrupo) & ";"
  End With
 Next
 ListadoAEnviar = "23" & ListadoAEnviar
 
 ' **************************************************************
 ' Envia el Listado
 ' **************************************************************
 EnviarPaqueteTCP ListadoAEnviar
                    
  
End Sub
Sub CambiarEstadoDeUsuario(IDUsuario As String, Estado As Integer)
Dim Contador As Integer
 
 ' **************************************************************
 ' Estados...
 ' 0. No Conectado
 ' 2. No Disponible
 ' 3. No Existe
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(IDUsuario)) Then
   If Estado <> 3 Then
     GrupoAmigo(Contador).EstadoDelAmigoEstado = Estado
     GrupoAmigo(Contador).EstadoDelAmigoTexto = ""
    Else
     GrupoAmigo(Contador).Existe = False
   End If
  End If
  CargarAmigos
 Next
 
End Sub
Sub CargarAmigos()
Dim RegrabarAmigos As Boolean
Dim Nodo As Node
Dim Contador, Contador2, ContadorTMP, TipoDeNodo, Cantidad, ColorNodo As Integer
Dim HijoDe, ImagenNodo, TextoNodo, UsuarioNombre, Existe, ImagenTMP As String
Dim Bandera, BanderaExpansion As Boolean
Dim PrimeraCarga As Boolean
   
  ' **************************************************************
  ' Arregla las Duplicidades...
  ' **************************************************************
  RegrabarAmigos = Varios.ArreglarAmigosRepetidos
  If RegrabarAmigos Then
   ' Si encontro diferencias manda a grabar, y espera el nuevo
   ' listado...
   Varios.GrabarListadoAmigos
   Exit Sub
  End If
  
  ' **************************************************************
  ' Borra el Listado de Amigos
  ' **************************************************************
  LimpiarListadoDeAmigos
    
  ' **************************************************************
  ' Carga el Nodo de Mensajes Pendientes
  ' **************************************************************
  ' * Usted Tiene (0) Mensajes Pendientes...
  Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "MensajesPendientes", MensajeRecurso(324), "MensajeFlash3", "MensajeFlash3")
  Nodo.Bold = False
  Nodo.ForeColor = vbBlack
  Nodo.BackColor = vbWhite
  Nodo.Expanded = True
  Varios.CambiarMensajesPendientes (Variables.CantidadDeMensajesPendientes)
  
  ' **************************************************************
  ' Carga el Primer Nodo el del Nombre del Usuario
  ' **************************************************************
  Cliente.ListadoDeAmigos.Enabled = True
  ' Define si la imagen debe ser de un Hombre o Una Mujer
  If UCase(Trim(Configuracion.Sexo)) = "F" Then
    ImagenTMP = "Mujer"
   Else
    ImagenTMP = "Hombre"
  End If
  UsuarioNombre = UCase(Mid$(Trim(Variables.Configuracion.IDAliasUsuario), 1, 1)) & Mid$(Trim(Variables.Configuracion.IDAliasUsuario), 2)
  ' Verifica si tiene o no Amigos
  If Variables.CantidadGrupoAmigo <> 0 Then
    ' Amigos de [ % ]...
    'Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(339) & Trim(UsuarioNombre) & MensajeRecurso(121) & Space(40), "Conectado", "Conectado")
    Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(339) & Trim(UsuarioNombre) & MensajeRecurso(121) & Space(40), ImagenTMP, ImagenTMP)
   Else
    ' ] No Posee Amigos...
    'Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(297) & Trim(UsuarioNombre) & MensajeRecurso(341), "Conectado", "Conectado")
    Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(, , "Usuario", MensajeRecurso(297) & Trim(UsuarioNombre) & MensajeRecurso(341), ImagenTMP, ImagenTMP)
  End If
  
  ' **************************************************************
  ' Seteos Especificos para el Nodo creado
  ' **************************************************************
  'Nodo.EnsureVisible
  Nodo.Bold = True
  Nodo.ForeColor = vbBlue
  Nodo.BackColor = vbWhite
  Nodo.Expanded = True
  
    
  ' **************************************************************
  ' Primero Carga Todos Los Grupos
  ' **************************************************************
  For Contador = 1 To Variables.CantidadGrupoAmigo
   With GrupoAmigo(Contador)
    If Trim(.NombreDelGrupo) <> "" Then
     ' **************************************************************
     ' Verifica que el Nodo que se Quiere Crear no Exista
     ' Actualmente...
     ' **************************************************************
     Bandera = False
     For ContadorTMP = 1 To Cliente.ListadoDeAmigos.Nodes.Count
      ' **************************************************************
      ' Primero se verifica que el nodo sea nodo
      ' **************************************************************
      If UCase(Mid$(Cliente.ListadoDeAmigos.Nodes(ContadorTMP).key, 1, 1)) = "G" Then
       ' **************************************************************
       ' Verifica que el nodo no este repetido
       ' **************************************************************
       If Mid$(Cliente.ListadoDeAmigos.Nodes(ContadorTMP).key, 2) = Trim(.NombreDelGrupo) Then
        Bandera = True
        Exit For
       End If
      End If
     Next
     ' **************************************************************
     If Bandera = False Then
      ' **************************************************************
      ' Crea el Nodo
      ' **************************************************************
      Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add("Usuario", tvwChild, "G" & Trim(.NombreDelGrupo), Trim(.NombreDelGrupo), "Grupo", "Grupo")
      ' **************************************************************
      ' Seteos Especificos para el Nodo creado
      ' **************************************************************
      ' Verifica si el nodo estaba NO expandido, caso en el cual,
      ' no lo expande... Por Default lo expande...
      BanderaExpansion = True
      If GrupoEstadoCantidad > 0 Then
       For Contador2 = 1 To GrupoEstadoCantidad
        If UCase(Trim(GrupoEstadoNombre(Contador2))) = UCase(Trim(.NombreDelGrupo)) Then
         ' El nodo no debe expandirse
         BanderaExpansion = False
         Exit For
        End If
       Next
      End If
      
      Nodo.Expanded = BanderaExpansion
      Nodo.Bold = True
      Nodo.BackColor = vbWhite
      Nodo.ForeColor = vbBlack
     End If
     ' **************************************************************
    End If
   End With
  Next
 
  ' **************************************************************
  ' Segundo Carga los Usuarios
  ' **************************************************************
  For Contador = 1 To CantidadGrupoAmigo
   With GrupoAmigo(Contador)
    If Trim(.IDNombreDelAmigo) <> "" Then
     ' **************************************************************
     ' Define a Donde se debe cargar el Usuario
     ' **************************************************************
     If Trim(.NombreDelGrupo) = "" Then
       HijoDe = "Usuario"
       TipoDeNodo = tvwChild
      Else
       HijoDe = "G" & Trim(.NombreDelGrupo)
       TipoDeNodo = tvwChild
     End If
     ' **************************************************************
     
     ' **************************************************************
     ' Define La Imagen a Poner y El Texto
     ' **************************************************************
     ImagenNodo = UCase(Trim(.Sexo)) & Trim(.EstadoDelAmigoEstado)
     TextoNodo = Trim(.IDNombreDelAmigo)
     ColorNodo = -1
     Select Case Trim(.EstadoDelAmigoEstado)
      Case "0" ' 0. No Conectado
       ' No Conectado...
       TextoNodo = TextoNodo & " (" & MensajeRecurso(287) & ")"
      Case "1" ' 1. Visible Normal
       ' Disponible (Normal)...
       TextoNodo = TextoNodo & " (" & MensajeRecurso(180) & ")"
      Case "2" ' 2. No Disponible
       ' No Disponible...
       TextoNodo = TextoNodo & " (" & MensajeRecurso(181) & ")"
      Case "3" ' 3. Custom
       TextoNodo = TextoNodo & " (" & Varios.ArreglarLenguaje(Trim(.EstadoDelAmigoTexto)) & ")"
     End Select
            
     ' **************************************************************
     ' Verifica si el Usuario Existe o no
     ' **************************************************************
     If Not .Existe Then
       ' Usuario Inexsitente...
       TextoNodo = Trim(.IDNombreDelAmigo) & " (" & MensajeRecurso(342) & ")"
       ImagenNodo = "UsuarioNoExiste"
       Existe = "0"
      Else
       Existe = "1"
     End If
     
     ' **************************************************************
     ' Atencion !!!!!!
     ' Atencion !!!!!!
     ' Ademas de Definirse la key del Nodo como
     ' U se establece su estado en e segundo caracter:
     ' Ej. U0 no Conectado, U2 No Disponible...
     ' **************************************************************
     Set Nodo = Cliente.ListadoDeAmigos.Nodes.Add(HijoDe, TipoDeNodo, "U" & Trim(.EstadoDelAmigoEstado) & Existe & Trim(.IDNombreDelAmigo), TextoNodo, ImagenNodo, ImagenNodo)
           
      
     ' **************************************************************
     ' Seteos Especificos para el Nodo creado
     ' **************************************************************
     'Nodo.EnsureVisible
     Nodo.Expanded = True
     Select Case Trim(.EstadoDelAmigoEstado)
      Case "0" ' 0. No Conectado
       Nodo.ForeColor = vbBlack
      Case "1" ' 1. Visible Normal
       Nodo.ForeColor = vbBlue
      Case "2" ' 2. No Disponible
       Nodo.ForeColor = vbBlack
      Case "3" ' 3. Custom
       Nodo.ForeColor = vbBlue
     End Select
     If Not .Existe Then
       Nodo.ForeColor = vbBlack
     End If
     Nodo.BackColor = vbWhite
    End If
   End With
  Next
  
  ' **************************************************************
  ' Pone la Cantidad de Usuario que tiene un Grupo Determinado
  ' **************************************************************
  For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
   If UCase(Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1)) = "G" Then
    Cantidad = Cliente.ListadoDeAmigos.Nodes(Contador).Children
    Select Case Cantidad
     Case 0
      '  (Ningún Amigo...)
      TextoNodo = MensajeRecurso(343)
     Case 1
      '  Amigo...)
      TextoNodo = " (" & CStr(Cantidad) & MensajeRecurso(344)
     Case Else
      '  Amigos...)
      TextoNodo = " (" & CStr(Cantidad) & MensajeRecurso(345)
    End Select
    Cliente.ListadoDeAmigos.Nodes(Contador).Text = Cliente.ListadoDeAmigos.Nodes(Contador).Text & TextoNodo
   End If
  Next
  
  ' **************************************************************
  ' Cambia en Los Formularios de Mensaje el Estado de los Amigos
  ' **************************************************************
  For Contador = 0 To Forms.Count - 1
   ' Busca el Formulario
   If Forms(Contador).FormularioNombre = "Mensajes" Or Forms(Contador).FormularioNombre = "DatosUsuario" Then
     Forms(Contador).PonerElEstadoDelUsuario
   End If
  Next
    
End Sub
Function CompletarCadena(Cadena As String, Largo As Integer, Lado As String, Caracter As String) As String
 Dim Contador As Integer
 Dim CadenaFinal As String
 
  ' **************************************************************
  ' Completa una Cadena del Lado Definido ([D]erecha o [I]zquierda
  ' con el [Caracter] especificado...
  ' **************************************************************
  CadenaFinal = Cadena
  For Contador = 1 To (Largo - Len(Cadena))
   If UCase(Lado) = "D" Then CadenaFinal = CadenaFinal & Caracter
   If UCase(Lado) = "I" Then CadenaFinal = Caracter & CadenaFinal
  Next
  
  ' **************************************************************
  ' Devuelve la Cadena
  ' **************************************************************
  CompletarCadena = CadenaFinal
 
End Function
Function MostrarMSGBox(Texto As String, Botones As Integer, Imagen As String, TituloVentana As String, Optional NoModal As Boolean) As Integer
Dim Tamanio1, Tamanio2, FormActual, Contador As Integer
Dim NuevoFormulario As New MensajesBox
Dim Bandera As Boolean

 'On Error GoTo SalirMostrarMSGBOX
 
 ' **************************************************************
 ' Verifica que no haya ningun Modal Abierto Caso en el cual
 ' Espera...
 ' **************************************************************
 Do Until Bandera = True
  DoEvents
  Bandera = True
  For Contador = 0 To Forms.Count - 1
   'DoEvents
   If Forms(Contador).FormularioNombre = "MensajesBox" Then
    If Forms(Contador).Modal Then Bandera = False
   End If
   If Forms(Contador).FormularioNombre = "IngresoBox" Then
    If Forms(Contador).Modal Then Bandera = False
   End If
  Next
 Loop
 
 ' **************************************************************
 ' Cierra todas las Ventanas de Menus que pudieren estar Abiertas
 ' **************************************************************
 CerrarVentanasDeMenus

 ' **************************************************************
 ' Carga y Muestra el Mensaje Solicitado por
 ' el Usuario
 ' **************************************************************
 Load NuevoFormulario
 FormActual = Forms.Count - 1
 Forms(FormActual).CargarDatosBox Texto, Botones, Imagen, TituloVentana
 ' Define si es modal o no...
 If NoModal = True Then
   Forms(FormActual).Modal = False
  Else
   Forms(FormActual).Modal = True
 End If
 
 ' **************************************************************
 ' Centra el Mensaje
 ' **************************************************************
 Tamanio1 = Forms(FormActual).TextoMensaje.HeighT
 Tamanio2 = 350 + Int((705 - Tamanio1) / 2)
 Forms(FormActual).TextoMensaje.Top = Tamanio2
 
 ' **************************************************************
 ' Muestra el Mensaje
 ' **************************************************************
 If NoModal = True Then
   Forms(FormActual).Show
  Else
   Forms(FormActual).Show vbModal
 End If
 
 ' **************************************************************
 ' Devuelve el Valor del Boton Clickeado
 ' **************************************************************
 MostrarMSGBox = Variables.RespuestaMensajeBox
 
SalirMostrarMSGBOX:
 Exit Function
 
End Function
Function MostrarInputBox(Texto As String, LargoIngreso As Integer, TituloVentana As String) As String

 ' **************************************************************
 ' Verifica que no haya ningun Modal Abierto Caso en el cual
 ' Espera...
 ' **************************************************************
 Dim Bandera As Boolean
 Dim Contador As Integer
 Do Until Bandera = True
  DoEvents
  Bandera = True
  For Contador = 0 To Forms.Count - 1
   DoEvents
   If Forms(Contador).FormularioNombre = "MensajesBox" Then
    If Forms(Contador).Modal Then Bandera = False
   End If
   If Forms(Contador).FormularioNombre = "IngresoBox" Then
    If Forms(Contador).Modal Then Bandera = False
   End If
  Next
 Loop
 
 ' **************************************************************
 ' Cierra todas las Ventanas de Menus que pudieren estar Abiertas
 ' **************************************************************
 CerrarVentanasDeMenus
 
 ' **************************************************************
 ' Carga y Muestra el Mensaje Solicitado al
 ' el Usuario
 ' **************************************************************
 Load IngresoBox
 IngresoBox.CargarDatosInputBox Texto, LargoIngreso, TituloVentana
  
 ' **************************************************************
 ' Muestra el Mensaje
 ' **************************************************************
 IngresoBox.Modal = True
 IngresoBox.Show vbModal
  
 ' **************************************************************
 ' Devuelve el Valor del Boton Clickeado
 ' **************************************************************
 MostrarInputBox = Variables.RespuestaIngresoBox
 
End Function
Sub VerDatosUsuario(Usuario As String, PermiteGrabar As Boolean)
Dim NuevoFormulario As New DatosUsuario
Dim NumeroFormulario  As Integer

 ' **************************************************************
 ' Genera y Carga el Nuevo Formulario
 ' **************************************************************
 Load NuevoFormulario
 NumeroFormulario = Forms.Count - 1
 ' Aca se completa el Cambio de Alias de Usuario, ya que esto
 ' se procesa despues del Evento On_Load del Formulario
 'Forms(NumeroFormulario).IDAliasUsuario = Usuario
 ' Carga los Datos compartidos del Formulario
 Forms(NumeroFormulario).AliasUsuario = Usuario
 Forms(NumeroFormulario).FormularioNombre = "DatosUsuario"
 Forms(NumeroFormulario).Refresco = False
 Forms(NumeroFormulario).CambioDeDatosUsuario = PermiteGrabar
 ' Cerrar
 Forms(NumeroFormulario).LabelOk = MensajeRecurso(128)
 
 If PermiteGrabar Then
  ' Activa los Controles para que Permiten Grabar los Datos
  Forms(NumeroFormulario).ShapeGrabar.Visible = True
  Forms(NumeroFormulario).LabelGrabar.Visible = True
  Forms(NumeroFormulario).BotonGrabar.Enabled = True
  Forms(NumeroFormulario).ApellidoYNombre.Locked = False
  Forms(NumeroFormulario).DireccionDeEmail.Locked = False
  Forms(NumeroFormulario).Edad.Locked = False
  Forms(NumeroFormulario).Sexo.Locked = False
  Forms(NumeroFormulario).UbicacionGeografica.Locked = False
  Forms(NumeroFormulario).Intencion.Locked = False
  Forms(NumeroFormulario).Humor.Locked = False
  Forms(NumeroFormulario).Ocupacion.Locked = False
  Forms(NumeroFormulario).Signo.Locked = False
  Forms(NumeroFormulario).EstadoCivil.Locked = False
  Forms(NumeroFormulario).Telefono.Locked = False
  Forms(NumeroFormulario).OtraInfo.Locked = False
  Forms(NumeroFormulario).FechaDeNacimiento.Locked = False
  ' Cancelar
  Forms(NumeroFormulario).LabelOk = MensajeRecurso(106)
  Forms(NumeroFormulario).ComboSexo.Enabled = True
  Forms(NumeroFormulario).ComboSigno.Enabled = True
  Forms(NumeroFormulario).ComboEstadoCivil.Enabled = True
 End If
 
 ' Muestra el Formulario
 Forms(NumeroFormulario).Show
 ' Pide los Datos del usuario
 Forms(NumeroFormulario).RefrescarDatos
 ' EnviarPaqueteTCP ("20" & CompletarCadena(Forms(NumeroFormulario).AliasUsuario, 16, "D", " "))
 
End Sub
Public Function PonerFocoEnVentana(Handle As Long)
Dim Contador As Integer

 ' **************************************************************
 ' Busca los Formularios con un Nombre Determinado
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
   ' Busca el Formulario con el nombre Necesario
  If Forms(Contador).hwnd = Handle Then
   ' Lo Encontro !!
   Forms(Contador).SeLlamoAlDesplegable = False
   Forms(Contador).Show
   Forms(Contador).ZOrder (0)
   Exit Function
  End If
 Next
 
  
End Function

Public Function DescargarVentanaHandle(Handle As Long)
Dim Contador As Integer

 ' **************************************************************
 ' Busca los Formularios con un Nombre Determinado
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
   ' Busca el Formulario con el nombre Necesario
  If Forms(Contador).hwnd = Handle Then
   ' Lo Encontro !!
   Forms(Contador).Hide
   Exit Function
  End If
 Next
 
  
End Function

Public Function BuscarFormularioNombre(NombreABuscar As String, Optional IDUsuario As String) As Integer
Dim Contador As Integer

 ' **************************************************************
 ' Busca los Formularios con un Nombre Determinado
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
  ' Busca el Formulario con el nombre Necesario
  DoEvents
  If Trim(UCase(Forms(Contador).FormularioNombre)) = Trim(UCase(NombreABuscar)) Then
   ' Lo Encontro !!
   If Trim(IDUsuario) <> "" Then
     If Trim(UCase(Forms(Contador).AliasUsuario)) = UCase(Trim(IDUsuario)) Then
      BuscarFormularioNombre = Contador
      Exit Function
     End If
    Else
     BuscarFormularioNombre = Contador
     Exit Function
   End If
  End If
 Next
 
 ' **************************************************
 
 ' ******************************************************
 ' No Lo Encontro !!
 ' **************************************************************
 BuscarFormularioNombre = -1
 
End Function
Public Sub LimpiarListadoDeAmigos()

  ' **************************************************************
  ' Esconde todos losMenus desplegables abiertos...
  ' **************************************************************
  CerrarMenusDecolgables
  
  ' **************************************************************
  ' Borra el Listado de Amigos
  ' **************************************************************
  Cliente.ListadoDeAmigos.Nodes.Clear

End Sub
Public Sub CerrarMenusDecolgables()
On Error GoTo cerrarMenusError:

  ' **************************************************************
  ' Cierra los Menus referentes de los Items del Listado
  ' **************************************************************
  Cliente.MenuMensajesPendientes.HideMenu
  Cliente.MenuClickUsuario.HideMenu
  Cliente.MenuClickAmigo.HideMenu
  Cliente.MenuClickGrupo.HideMenu
  Exit Sub
  
cerrarMenusError:
 Resume Next
  
End Sub
Public Function MensajeRecursoReal(MensajeNumero As Integer) As String

 ' **************************************************************
 ' Baja el Mensaje del Resource File segun el Idioma
 ' **************************************************************
 MensajeRecursoReal = LoadResString(MensajeNumero)

End Function

Public Function MensajeRecurso(MensajeNumero As Integer) As String
Dim LenguajeID As Integer
Dim MensajeTMP As String

  LenguajeID = 0
  
 ' **************************************************************
 ' Define el Numero segun el Lenguaje...
 ' **************************************************************
 Select Case Trim(UCase(Variables.LenguajeActual))
  Case "ESPAÑOL":
   LenguajeID = 0
  Case "SPANISH":
   LenguajeID = 0
  Case "INGLES":
   LenguajeID = 1000
  Case "ENGLISH":
   LenguajeID = 1000
 End Select
 
 ' **************************************************************
 ' Baja el Mensaje del Resource File segun el Idioma
 ' **************************************************************
 MensajeRecurso = LoadResString(MensajeNumero + LenguajeID)
 
End Function
Public Function ArreglarLenguaje(TextoMensaje As String) As String
Dim MensajeTemp, MensajeTemp2, Mensaje As String
Dim Contador As Integer
  
  ' **************************************************************
  ' Arregla las diferencias de Lenguaje...
  ' **************************************************************
  MensajeTemp = UCase(Trim(TextoMensaje))
  MensajeTemp2 = ""
  
  ' Español...
  If UCase(Trim(Variables.LenguajeActual)) = "ESPAÑOL" Or UCase(Trim(Variables.LenguajeActual)) = "SPANISH" Then
   For Contador = 0 To 3
    Mensaje = UCase(Trim(MensajeRecursoReal(1180 + Contador)))
    If MensajeTemp = Mensaje Then
     MensajeTemp2 = Trim(MensajeRecursoReal(180 + Contador))
    End If
   Next
  End If
  
  ' Ingles...
  If UCase(Trim(Variables.LenguajeActual)) = "ENGLISH" Or UCase(Trim(Variables.LenguajeActual)) = "INGLES" Then
   For Contador = 0 To 3
    Mensaje = UCase(Trim(MensajeRecursoReal(180 + Contador)))
    If MensajeTemp = UCase(Trim(MensajeRecursoReal(180 + Contador))) Then
     MensajeTemp2 = Trim(MensajeRecursoReal(1180 + Contador))
    End If
   Next
  End If
  
  ' **************************************************************
  ' Verifica que Debe Devolver...
  ' **************************************************************
  If MensajeTemp2 = "" Then
    ArreglarLenguaje = TextoMensaje
   Else
    ArreglarLenguaje = MensajeTemp2
  End If
  
End Function
Public Function SolicitarAgregarAmigo(AliasAmigo As String) As Integer
Dim TiempoInicial As Date
Dim SegundosTranscurridos As Integer

 ' **************************************************************
 ' Enviar la Query al Servidor Sobre el Amigo
 ' **************************************************************
 EnviarPaqueteTCP ("24" & CompletarCadena(CStr(AliasAmigo), 16, "D", " "))
 
 ' **************************************************************
 ' Espera 5 segundos por el OK del Usuario
 ' **************************************************************
 Variables.RespuestaAmigoACrearResultado = -1
 TiempoInicial = Time
 Do
  DoEvents
  If Variables.RespuestaAmigoACrearResultado <> -1 And UCase(Trim(Variables.RespuestaAmigoACrearAlias)) = UCase(Trim(AliasAmigo)) Then Exit Do
  SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 
 ' **************************************************************
 ' Devuelve el Resultado...
 ' **************************************************************
 SolicitarAgregarAmigo = Variables.RespuestaAmigoACrearResultado
 
End Function

Public Function BuscarVentanaMensajeOffLine(Usuario As String) As Integer
Dim Contador As Integer
   
 ' **************************************************************
 ' Busca Formulario OffLine
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
  If UCase(Trim(Forms(Contador).FormularioNombre)) = UCase("MensajeOffline") And UCase(Trim(Forms(Contador).AliasUsuario)) = UCase(Trim(Usuario)) Then
   BuscarVentanaMensajeOffLine = Contador
   Exit Function
  End If
 Next
 
 BuscarVentanaMensajeOffLine = 0
 
End Function
Public Function EsAmigo(AmigoID As String) As Boolean
Dim Contador As Integer

 ' **************************************************************
 ' Si es amigo manda un TRUE
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(AmigoID)) Then
   EsAmigo = True
   Exit Function
  End If
 Next
  
 ' **************************************************************
 ' No Lo Encontro...
 ' **************************************************************
 EsAmigo = False
 
End Function
Function EnviarBorradousuario(ResponsableNombre As String, ResponsableVentanaID As String, BorrarNombre As String, BorrarVentanaID As String)
Dim ComandoAdicional As String

   ComandoAdicional = "43" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
                    "M" & "01" & _
                    CompletarCadena(Trim(ResponsableNombre), 16, "D", " ") & _
                    CompletarCadena(Trim(ResponsableVentanaID), 10, "I", "0") & "5" & _
                    CompletarCadena(Trim(BorrarVentanaID), 10, "I", "0") & CompletarCadena(Trim(BorrarNombre), 16, "D", " ")
   EnviarPaqueteTCP ComandoAdicional

End Function


