Attribute VB_Name = "Inicializar"
Option Explicit
Public Function PosicionCliente(Comando As String, Optional PosicionX As Integer, Optional PosicionY As Integer) As POINTAPI
Dim PosicionVentana As POINTAPI

 Select Case UCase(Comando)
  Case "GRABAR"
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionX", CStr(PosicionX)
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionY", CStr(PosicionY)
  Case "LEER"
   If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionX")) Then
     PosicionVentana.X = -10000
    Else
     PosicionVentana.X = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionX")
   End If
   If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionY")) Then
     PosicionVentana.Y = -10000
    Else
     PosicionVentana.Y = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PosicionY")
   End If
   PosicionCliente = PosicionVentana
 End Select
 
End Function
Sub Main()

 ' **********************************************************************
 ' Inicializar el Sistema...
 ' **********************************************************************
 Inicializar.CargarConfiguracionArchivo "NO"

 ' **************************************************************
 ' Carga la Configuracion del Usuario
 ' **************************************************************
 With Configuracion
  .Logueado = False
  .TimeOutLogueo = 10
  .TimeOutGeneral = 7
  .TimeOutMultiChat = 10
  .NombreDelSistema = "EIM"
  .VersionDelSistema = "1.00"
  ' **********************************************************************
  ' Levantar el Lenguaje
  ' **********************************************************************
  '.Lenguaje = Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "Lenguaje"))
  'If .Lenguaje = "" Then .Lenguaje = "English"
  Variables.LenguajeActual = .Lenguaje
  ' Versión
  .TituloVentanas = Trim(Trim(Configuracion.NombreDelSistema) & " " & _
                                MensajeRecurso(313) & Trim(Configuracion.VersionDelSistema))
  '.TituloVentanas = Trim(Trim(Configuracion.NombreDelSistema) & " " & Trim(Configuracion.VersionDelSistema))
 End With
 
 ' **********************************************************************
 ' Carga la variable de ID Mensaje
 ' **********************************************************************
 IDMensaje = 0
 
 ' **********************************************************************
 ' Define los Colores...
 ' **********************************************************************
 ''ShapesBackColor= vbGreen ' &HE0E0E0
 ''ShapesBorderColor = vbYellow ' &HFF8080
 ''FontLabelColor = vbMagenta ' vbblack
 ''FontHipervinculoColor = vbCyan ' vbblue
 ''FontTituloVentana = vbWhite ' &H00400000&
 ''FontBotonesColor = vbRed ' vbblack
 ''FontMenuDescolgable = vbRed ' vbblack
 ''FontMenuDescolgableAbierto = vbYellow ' vbblack
 ''FontFOndoMenuDescolgable = vbBlue ' &HE0E0E0
 ''FontMenuDescolgableAbiertoFranjas = vbRed ' &HE0E0E0
 ''FontMenuDescolgableAbiertoFondoFormulario = vbCyan ' VBWhite
 ''FontMenuDescolgableAbiertoLineaOscura = vbBlue '&H00808080&
 ''FontMenuDescolgableAbiertoHighLightLetra = vbCyan ' vbwithe
 ''FontMenuDescolgableAbiertoFont = vbYellow ' vbblack
 ''FontMenuDescolgableAbiertoHighLightFondo = vbWhite ' vbblue
 ''FontMenuDescolgableAbiertoHighLightBorde = vbGreen ' &H00FF0000&
 ''FontMenuDescolgableAbiertoMostrarGrafico = False ' true
 ''FontFormClienteTextosDeEstado = vbGreen ' vbblack
 ''FontFormClienteTimer = vbRed  ' VBblack
 ''FontCambioDeEstado1 = vbBlack
 ''FontCambioDeEstado2 = vbWhite
 ''FontCambiarLetraNormalBack = vbRed ' vbwith
 ''FontCambiarLetraNormalBorder = vbGreen ' vbblack
 ''FontCambiarLetraHighBack = vbYellow '&HFF8080
 ''FontCambiarLetraHighBorder = vbBlue ' &HFF0000
 ''FontCambiarLetraLetra = vbCyan 'vbblack
 ''FontMensajeBotonApretadoFondo = vbRed  '
 ''FontMensajeBotonColorLetra = vbWhite 'vbBlack


 ' **
 ShapesBackColor = &HE0E0E0
 ShapesBorderColor = &HFF8080
 FontLabelColor = vbBlack
 FontHipervinculoColor = vbBlue
 FontTituloVentana = &H400000
 FontBotonesColor = vbBlack
 FontMenuDescolgable = vbBlack
 FontFOndoMenuDescolgable = &HE0E0E0
 FontMenuDescolgableAbierto = vbBlack
 FontMenuDescolgableAbiertoFranjas = &HE0E0E0
 FontMenuDescolgableAbiertoFondoFormulario = vbWhite
 FontMenuDescolgableAbiertoLineaOscura = &H808080
 FontMenuDescolgableAbiertoFont = vbBlack
 FontMenuDescolgableAbiertoHighLightLetra = vbWhite
 FontMenuDescolgableAbiertoHighLightFondo = &HFF8080
 FontMenuDescolgableAbiertoMostrarGrafico = True
 FontMenuDescolgableAbiertoHighLightBorde = &HFF0000
 FontFormClienteTextosDeEstado = vbBlack
 FontFormClienteTimer = vbBlack
 FontCambioDeEstado1 = vbBlack ' Ver estos colores...
 FontCambioDeEstado2 = vbWhite ' Ver estos colores...
 FontCambiarLetraNormalBack = vbWhite
 FontCambiarLetraNormalBorder = vbBlack
 FontCambiarLetraHighBack = &HFF8080
 FontCambiarLetraHighBorder = &HFF0000
 FontCambiarLetraLetra = vbBlack
 FontMensajeBotonApretadoFondo = &H808080
 FontMensajeBotonColorLetra = vbBlack
 
 ' **********************************************************************
 ' Cargar los Sonidos
 ' **********************************************************************
 BufferSonido1 = LoadResData(1, "Sonidos")
 BufferSonido2 = LoadResData(2, "Sonidos")
 BufferSonido3 = LoadResData(3, "Sonidos")
 BufferSonido4 = LoadResData(4, "Sonidos")
 BufferSonido5 = LoadResData(5, "Sonidos")
 
 ' **************************************************************
 ' Ejecuta el Sonido de Bienvenida
 ' **************************************************************
 EjecutarSonido 5, "SI"
 
 ' **********************************************************************
 ' Carga la Presentacion
 ' **********************************************************************
 Load Presentacion
 Presentacion.Show vbModal
  
 ' **********************************************************************
 ' Carga el Cliente
 ' **********************************************************************
 Load Cliente
 ' Si no tiene que estar minimizado Muestra el Formulario
 If Not Configuracion.CargarMinimizado Then
  DefinirPosicionDeCliente True
 End If
 
End Sub
Sub GrabarEstadoDeGrupos()
Dim Cantidad, Contador As Integer

 ' **********************************************************************
 ' Solo graba si esta logueado
 ' **********************************************************************
 If Configuracion.Logueado <> 3 Then Exit Sub
 
 ' **********************************************************************
 ' Lleva a nada los Anteriores...
 ' **********************************************************************
 If IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad")) Then
  Cantidad = CInt(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad"))
  If Cantidad <> 0 Then
   For Contador = 1 To Cantidad
    GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Nombre" & Contador, ""
   Next
  End If
 End If
 
 ' **********************************************************************
 ' Graba La Cantidad de EstadosGrupos
 ' **********************************************************************
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad", CStr(Variables.GrupoEstadoCantidad)
 
 ' **********************************************************************
 ' Graba Los Nombres de los Grupos minimizados
 ' **********************************************************************
 For Contador = 1 To Variables.GrupoEstadoCantidad
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Nombre" & Contador, Trim(Variables.GrupoEstadoNombre(Contador))
 Next

End Sub
Sub CargarEstadoDeGrupos()
Dim Contador, Cantidad As Integer

 ' **********************************************************************
 ' Pone a cero la cantidad de Grupos...
 ' **********************************************************************
 GrupoEstadoCantidad = 0
 
 ' **********************************************************************
 ' Valida que existan Grupos Grabados
 ' **********************************************************************
 If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad")) Then
  Exit Sub
 End If
 
 ' **********************************************************************
 ' Carga La Cantidad
 ' **********************************************************************
 Cantidad = CInt(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad"))
 If Cantidad = 0 Then Exit Sub
 
 ' **********************************************************************
 ' Carga y Define la Cantidad Grupos Minimizados
 ' **********************************************************************
 For Contador = 1 To Cantidad
  GrupoEstadoCantidad = GrupoEstadoCantidad + 1
  ReDim Preserve GrupoEstadoNombre(GrupoEstadoCantidad)
  GrupoEstadoNombre(GrupoEstadoCantidad) = (CStr(Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\EstadoGrupos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Nombre" & Contador))))
 Next
 
End Sub
Sub CargarBloqueUsuarios()
Dim Contador, Cantidad As Integer

 ' **********************************************************************
 ' Valida que existan usuarios Bloqueados, si no los ahi sale...
 ' **********************************************************************
 If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Bloqueos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad")) Then
  Variables.UsuarioBloqueadoCantidad = 0
  Exit Sub
 End If
 
 ' **********************************************************************
 ' Carga y Define la Cantidad de usuarios Bloqueados
 ' **********************************************************************
 Variables.UsuarioBloqueadoCantidad = (CInt(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Bloqueos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad")))
 ReDim Variables.UsuarioBloqueadoNombres(Variables.UsuarioBloqueadoCantidad)
 
 ' **********************************************************************
 ' Carga los Usuarios Bloqueado
 ' **********************************************************************
 For Contador = 1 To Variables.UsuarioBloqueadoCantidad
  Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario = CStr(Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Bloqueos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Nombre" & Contador)))
 Next
 
End Sub
Sub GrabarBloqueUsuarios()
Dim Contador As Integer
Dim UsuarioBloqueadoTemp, Archivo, NombreArchivo As String

 ' **********************************************************************
 ' Graba La Cantidad de Bloqueos
 ' **********************************************************************
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Bloqueos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Cantidad", CStr(Variables.UsuarioBloqueadoCantidad)
 
 ' **********************************************************************
 ' Graba Los Bloqueos
 ' **********************************************************************
 For Contador = 1 To Variables.UsuarioBloqueadoCantidad
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Bloqueos\" & Trim(UCase(Configuracion.IDAliasUsuario)), "Nombre" & Contador, Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)
 Next

End Sub
Public Function UltimoEstado(Operacion As String, Usuario As String)
Dim UltimoEstadoNumero, UltimoEstadoTexto As String

 ' **********************************************************************
 ' Las operaciones son Grabar y Leer
 ' **********************************************************************
 Select Case UCase(Operacion)
  Case "GRABAR"
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoNumero", CStr(Configuracion.EstadoDelUsuario)
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoTexto", CStr(Configuracion.EstadoActualTexto)
  Case "LEER"
   UltimoEstadoNumero = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoNumero")
   ' Verifica que sea numeric, si lo es lo devuelve, sino envia -1
   If IsNumeric(UltimoEstadoNumero) Then
     Variables.UltimoEstadoDelUsuario.EstadoNumero = CInt(UltimoEstadoNumero)
    Else
     Variables.UltimoEstadoDelUsuario.EstadoNumero = -1
   End If
   Variables.UltimoEstadoDelUsuario.Estadotexto = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoTexto")
  Case "PONER"
   UltimoEstadoNumero = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoNumero")
   UltimoEstadoTexto = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & Trim(UCase(Usuario)), "UltimoEstadoTexto")
   ' Si no es numerico Sale
   If Not IsNumeric(UltimoEstadoNumero) Then Exit Function
   If IsNull(UltimoEstadoTexto) Then UltimoEstadoTexto = ""
   ' Pone la info
   Configuracion.EstadoDelUsuario = UltimoEstadoNumero
   Configuracion.EstadoActualTexto = UltimoEstadoTexto
 End Select

End Function
Sub CargarConfiguracionArchivo(Optional Todo As String)

 ' **********************************************************************
 ' Pasar la Configuracion Para Levantda...
 ' **********************************************************************
 With Variables.Configuracion
  ' **********************************************************************
  ' Validar el Lenguaje
  ' **********************************************************************
  .Lenguaje = Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "Lenguaje"))
  If .Lenguaje = "" Then .Lenguaje = "English"
  ' **********************************************************************
  ' El ID Alias no es necesario Validarlo
  ' **********************************************************************
  .IDAliasUsuario = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "IDAliasUsuario")
  ' **********************************************************************
  ' Validar el Port TCP
  ' **********************************************************************
  If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PortTCP")) Then
    .PortTCP = 24157
   Else
    .PortTCP = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "PortTCP")
  End If
  If .PortTCP = 0 Then .PortTCP = 24157
  ' **********************************************************************
  ' Valida el Nombre del Servidor de Conecccion
  ' **********************************************************************
  .Servidor = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "Servidor")
  If Trim(.Servidor) = "" Or Mid$(.Servidor, 1, 1) = Chr$(0) Then .Servidor = "127.0.0.1"
  ' **********************************************************************
  ' Recordar Password Estado
  ' **********************************************************************
  If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "RecordarPasswordEstado") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "RecordarPasswordEstado") <> "No" Then
     .RecordarPasswordEstado = False
     .RecordarPasswordPassword = ""
    Else
     If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "RecordarPasswordEstado") = "Si" Then
       .RecordarPasswordEstado = True
      Else
       .RecordarPasswordEstado = False
       .RecordarPasswordPassword = ""
     End If
  End If
  ' **********************************************************************
  ' Valida el Recordar Password Password
  ' **********************************************************************
  .RecordarPasswordPassword = Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "RecordarPasswordPassword"))
  ' **********************************************************************
  ' Logueo Automatico
  ' **********************************************************************
  If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "LogueoAutomatico") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "LogueoAutomatico") <> "No" Then
     .LogueoAutomatico = False
    Else
     If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "LogueoAutomatico") = "Si" Then
       .LogueoAutomatico = True
      Else
       .LogueoAutomatico = False
     End If
  End If
  ' **********************************************************************
  ' Arranque Con Windows
  ' **********************************************************************
  If Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "EIM")) <> "" Then
    .ArranqueConWindows = True
    ' Verifica que este bien el Path, sino graba el nuevo path
    If Trim(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "EIM")) <> App.Path & "\" & App.EXEName & ".exe" Then
     GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "EIM", App.Path & "\" & App.EXEName & ".exe"
    End If
   Else
    .ArranqueConWindows = False
  End If
  ' **********************************************************************
  ' Cargar Minimizado
  ' **********************************************************************
  If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "CargarMinimizado") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "CargarMinimizado") <> "No" Then
     .CargarMinimizado = False
    Else
     If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "CargarMinimizado") = "Si" Then
       .CargarMinimizado = True
      Else
       .CargarMinimizado = False
     End If
  End If
 
  ' **********************************************************************
  ' Valida la configuracion de Sonido
  ' **********************************************************************
  If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "SonidoActivado") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "SonidoActivado") <> "No" Then
     .SonidoActivado = True
    Else
     If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion", "SonidoActivado") = "Si" Then
       .SonidoActivado = True
      Else
       .SonidoActivado = False
     End If
  End If
 
' **********************************************************************
' **********************************************************************
  ' Estos son Cargados segun el Usuario...
  If Trim(.IDAliasUsuario) <> "" And UCase(Trim(Todo)) <> "NO" Then
   ' **********************************************************************
   ' Validar Para Pasar a Enseguida Vuelvo
   ' **********************************************************************
   If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "TiempoParaPasaraInactivo")) Then
     .TiempoParaPasaraInactivo = 5
    Else
     .TiempoParaPasaraInactivo = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "TiempoParaPasaraInactivo")
   End If
   If .TiempoParaPasaraInactivo <= 0 Then .TiempoParaPasaraInactivo = 1
  
   ' **********************************************************************
   ' Validar Tiempo de Refresco
   ' **********************************************************************
   If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "TiempoDeRefrescoAmigos")) Then
     .TiempoDeRefrescoAmigos = 1
    Else
     .TiempoDeRefrescoAmigos = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "TiempoDeRefrescoAmigos")
   End If
   If .TiempoDeRefrescoAmigos <= 0 Or .TiempoDeRefrescoAmigos >= 10 Then .TiempoDeRefrescoAmigos = 1
  
   ' **********************************************************************
   ' Avisar Sobre Mensajes Pendientes
   ' **********************************************************************
   If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "AvisarMensajesPendientes") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "AvisarMensajesPendientes") <> "No" Then
      .AvisarMensajesPendientes = True
     Else
      If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "AvisarMensajesPendientes") = "Si" Then
        .AvisarMensajesPendientes = True
       Else
        .AvisarMensajesPendientes = False
      End If
   End If
   ' **********************************************************************
   ' Valida el Informar cambio de Estado
   ' **********************************************************************
   If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "InformarCambiosDeEstado") <> "Si" And LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "InformarCambiosDeEstado") <> "No" Then
      .InformarCambiosDeEstado = True
     Else
      If LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "InformarCambiosDeEstado") = "Si" Then
        .InformarCambiosDeEstado = True
       Else
        .InformarCambiosDeEstado = False
      End If
   End If
   ' **********************************************************************
   ' Valida el Directorio de Download
   ' **********************************************************************
   .DirectorioDownload = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "DirectorioDownload")
   If Mid$(.DirectorioDownload, 1, 1) = Chr$(0) Then .DirectorioDownload = ""
   ' **********************************************************************
   ' Valida el Directorio de Upload
   ' **********************************************************************
   .DirectorioUpload = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "DirectorioUpload")
   If Mid$(.DirectorioUpload, 1, 1) = Chr$(0) Then .DirectorioUpload = ""
   ' **********************************************************************
   ' Valida el Font Estandar
   ' **********************************************************************
   .FontEstandarNombre = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "FontEstandarNombre")
   If Mid$(.FontEstandarNombre, 1, 1) = Chr$(0) Or Trim(.FontEstandarNombre) = "" Then .FontEstandarNombre = "Arial"
   If Not IsNumeric(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "FontEstandarTamano")) Then
     .FontEstandarTamano = 8
    Else
     If CInt(LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "FontEstandarTamano")) = 0 Then
       .FontEstandarTamano = 8
      Else
       .FontEstandarTamano = LeerRegistry(HKEY_LOCAL_MACHINE, "Software\EIM\Configuracion\" & UCase(Trim(.IDAliasUsuario)), "FontEstandarTamano")
     End If
   End If
  End If
  
  ' **********************************************************************
  ' Graba la Configuracion...
  ' **********************************************************************
  ' Graba lo que se pidio Leer...
  GrabarConfiguracion Todo
 End With
  
End Sub
Sub GrabarConfiguracion(Optional Todo As String)
 
 ' **********************************************************************
 ' Graba la configuracion en Registry
 ' **********************************************************************
 If Trim(Configuracion.IDAliasUsuario) <> "" And UCase(Trim(Todo)) <> "NO" Then
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "TiempoDeRefrescoAmigos", Trim(CStr(Variables.Configuracion.TiempoDeRefrescoAmigos))
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "TiempoParaPasaraInactivo", Trim(CStr(Variables.Configuracion.TiempoParaPasaraInactivo))
  If Variables.Configuracion.AvisarMensajesPendientes Then
    GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "AvisarMensajesPendientes", "Si"
   Else
    GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "AvisarMensajesPendientes", "No"
  End If
  If Variables.Configuracion.InformarCambiosDeEstado Then
    GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "InformarCambiosDeEstado", "Si"
   Else
    GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "InformarCambiosDeEstado", "No"
  End If
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "DirectorioDownload", Trim(Variables.Configuracion.DirectorioDownload)
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "DirectorioUpload", Trim(Variables.Configuracion.DirectorioUpload)
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "FontEstandarNombre", Trim(Variables.Configuracion.FontEstandarNombre)
  GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion\" & Trim(UCase(Configuracion.IDAliasUsuario)), "FontEstandarTamano", Trim(Variables.Configuracion.FontEstandarTamano)
 End If
 
 If Variables.Configuracion.SonidoActivado Then
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "SonidoActivado", "Si"
  Else
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "SonidoActivado", "No"
 End If
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "Lenguaje", Trim(CStr(Variables.Configuracion.Lenguaje))
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "IDAliasUsuario", Trim(Variables.Configuracion.IDAliasUsuario)
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "PortTCP", Trim(CStr(Variables.Configuracion.PortTCP))
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "Servidor", Trim(Variables.Configuracion.Servidor)
 If Variables.Configuracion.RecordarPasswordEstado Then
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "RecordarPasswordEstado", "Si"
  Else
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "RecordarPasswordEstado", "No"
 End If
 GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "RecordarPasswordPassword", Trim(Configuracion.RecordarPasswordPassword)
 If Variables.Configuracion.LogueoAutomatico Then
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "LogueoAutomatico", "Si"
  Else
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "LogueoAutomatico", "No"
 End If
 If Variables.Configuracion.ArranqueConWindows Then
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "EIM", App.Path & "\" & App.EXEName & ".exe"
  Else
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "EIM", ""
 End If
 If Variables.Configuracion.CargarMinimizado Then
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "CargarMinimizado", "Si"
  Else
   GrabarRegistry HKEY_LOCAL_MACHINE, "Software\Eim\Configuracion", "CargarMinimizado", "No"
 End If
  
End Sub
Sub InicializarSistema()

 ' **************************************************************
 ' Carga la Configuracion del Archivo...
 ' **************************************************************
 CargarConfiguracionArchivo "NO"
 CargarBloqueUsuarios

 ' **********************************************************************
 ' Cargar las Fuentes
 ' **********************************************************************
 Varios.CargarFuentes
 
 ' **********************************************************************
 ' Definir Estado del Sonido
 ' **********************************************************************
 DefinirEstadoSonido
 
 ' **************************************************************
 ' Inicializa los Parametros del Socket
 ' **************************************************************
 Cliente.TCPSocket.RemoteHost = Configuracion.Servidor
 Cliente.TCPSocket.RemotePort = Configuracion.PortTCP
 Cliente.TCPSocket.Protocol = sckTCPProtocol
 
 
 ' **************************************************************
 ' Define como Empieza el Contador de Paquetes Partidos
 ' **************************************************************
 CantidadDePaquetesPartidos = 0
  
 ' **************************************************************
 ' Define que no se aviso sobre mensaje pendientes
 ' **************************************************************
 Variables.UltimoMensajePendienteAviso = False
  
 ' **************************************************************
 ' Define que no se paso a Enseguida Vuelvo en forma automatica
 ' **************************************************************
 Variables.PasoAEnseguidaVuelvo = False
  
 ' **************************************************************
 ' Carga el Formulario de Cliente para que no de errro el Cambio
 ' de Estado...
 ' **************************************************************
 Load Cliente
 Cliente.CargarLosMenus
 CambiarColorMenus.CambiarColorTreeList (RGB(255, 255, 255))
 
 ' **************************************************************
 ' Inicializa el Tray Icon
 ' **************************************************************
 Set Cliente.SysIcon = New CSystrayIcon
 Dim Nombre As String
 Nombre = Trim(Configuracion.TituloVentanas)
 Cliente.SysIcon.Initialize Cliente.hwnd, Cliente.IconoTrayDesConectado.Picture, Nombre
 Cliente.SysIcon.ShowIcon
 ' ************************************************

 ' **************************************************************
 ' Define el Estado del Sistema en NO-LOGUEADO
 ' **************************************************************
 SocketTCP.CambiarEstadoDelCliente (0)
 
 ' **************************************************************
 ' Por Default el Estado del Usuario es 1 (Visible)
 ' **************************************************************
 Configuracion.EstadoDelUsuario = 1
 Configuracion.EstadoActualTexto = ""

 ' **************************************************************
 ' Define el Refresco del Listado de Amigos
 ' **************************************************************
 Cliente.RefrescoAmigos = 65000 * Configuracion.TiempoDeRefrescoAmigos
 
 
End Sub

Public Function DefinirEstadoSonido()

 ' **************************************************************
 ' Define la Imagen del Sonido
 ' **************************************************************
 If Configuracion.SonidoActivado Then
   Cliente.SonidoSeteo.Picture = Cliente.ImagenesSonido.ListImages("ConSonido").Picture
   Cliente.SonidoSeteo.ToolTipText = MensajeRecurso(314)
  Else
   Cliente.SonidoSeteo.Picture = Cliente.ImagenesSonido.ListImages("SinSonido").Picture
   Cliente.SonidoSeteo.ToolTipText = MensajeRecurso(315)
 End If

End Function
