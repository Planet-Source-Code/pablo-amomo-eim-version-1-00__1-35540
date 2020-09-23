Attribute VB_Name = "Varios"
Option Explicit
Sub IniciarIconTray()

 ' **************************************************************
 ' Inicializa el Tray Icon
 ' **************************************************************
 'Set Server.SysIcon = New CSystrayIcon
 'Dim Nombre As String
 'Nombre = Trim(Configuracion.TituloVentanas)
 'Server.SysIcon.Initialize Server.hwnd, Server.EstadoSistema.ListImages("IconoSistema").Picture, Nombre
 'Server.SysIcon.ShowIcon
 ' ************************************************

End Sub

Function GenerarPassword(Caracteres As Integer) As String
Dim contador, Numero As Integer
Dim Password As String

 For contador = 1 To Caracteres
  Randomize
  Numero = Int((122 - 97 + 1) * Rnd + 97)
  Password = Password & Chr(Numero)
 Next
 
 GenerarPassword = Password
 
End Function
Function FechaActualFormatoOffLine() As String
Dim Fecha As String

 ' **************************************************************
 ' Arma la Fecha con el  Formato:
 ' HH:MM:SS_DD/MM/AAAA
 ' **************************************************************
 Fecha = ""
 Fecha = Fecha & Format(Time, "hh:mm:ss") ' Hora
 Fecha = Fecha & "_"
 Fecha = Fecha & Format(Date, "DD/MM/YYYY") ' Fecha
 FechaActualFormatoOffLine = Fecha
 
End Function

Function FechaActual() As String
Dim Fecha As String

 ' **************************************************************
 ' Arma la Fecha con el  Formato:
 ' DD/MM/AAAA - HH:MM:SS AM/PM
 ' **************************************************************
 Fecha = ""
 Fecha = Fecha & Format(Date, "DD/MM/YYYY") ' Fecha
 Fecha = Fecha & " - "
 Fecha = Fecha & Format(Time, "hh:mm:ss AMPM") ' Hora
 FechaActual = Fecha
 
End Function
Sub CargarPantallasAutomaticas()
 
 ' **************************************************************
 ' Carga todas las pantallas que deben refrescarse
 ' en forma automatica
 ' **************************************************************
 CargarPantallaDeConfiguracion
 CargarListadoUsuariosActuales
 CargarListadoDeUsuarioRegistrado
 
End Sub
Function BuscaElPrimerUserIDDisponible() As Integer
Dim contador As Integer

 ' **************************************************************
 ' Ubica el Primer ID libre para cargar un nuevo
 ' Usuario
 ' **************************************************************
 For contador = 1 To Configuracion.CantidadDeUsuarios
  If Trim(CStr(Usuarios(contador).IDAliasUsuario)) = "" Then
   BuscaElPrimerUserIDDisponible = contador
   Exit For
  End If
 Next
 
 BuscaElPrimerUserIDDisponible = 0
 
End Function
Sub CargarListadoDeUsuarioRegistrado()
Dim contador As Integer
Dim AliasUsuario As String

 ' **************************************************************
 ' Carga los Usuarios Definidos en el Sistema
 ' **************************************************************
 Server.UsuariosRegistrados.Clear
 For contador = 1 To Configuracion.CantidadDeUsuarios
  If Trim(CStr(Usuarios(contador).IDAliasUsuario)) <> "" Then
   Server.UsuariosRegistrados.AddItem (Trim(Usuarios(contador).IDAliasUsuario) & " - [" & Trim(Usuarios(contador).ApellidoYNombre) & "]")
  End If
 Next
 
 ' **************************************************************
 ' No hay Usuario en el Listado...
 ' **************************************************************
 If Server.UsuariosRegistrados.ListCount = 0 Then
  Server.BlanquearCamposUserManager
  Exit Sub
 End If

 ' **************************************************************
 ' Se posiciona en el Primer Usuario...
 ' **************************************************************
 Server.UsuariosRegistrados.ListIndex = 0
 ' Separa el IDDeUsuario
 contador = InStr(1, Server.UsuariosRegistrados.Text, "- [")
 If contador > 0 Then ' Separa el IDAlias del Usuario...
  AliasUsuario = Mid$(Server.UsuariosRegistrados.Text, 1, contador - 1)
 End If
 Server.CargarUsuario_UserManager (AliasUsuario)

End Sub
Sub CargarPantallaDeConfiguracion()

 ' **************************************************************
 ' Carga los seteos de Configuracion en la pantalla de Configuración
 ' **************************************************************
 Server.UbicacionBaseDeDatos = Trim(Configuracion.UbicacionBaseDeDatos)
 Server.NombreDeLaBaseDeDatos = Trim(Configuracion.NombreDeLaBaseDeDatos)
 Server.UsuariosSoportados = Configuracion.UsuariosSoportados
 Server.PortTCP = Configuracion.PortTCP
 If Configuracion.PermitirCrear Then
   Server.PermitirCrearUsuarios = 1
  Else
   Server.PermitirCrearUsuarios = 0
 End If
 Server.DireccionSMTP = Trim(Configuracion.DireccionIPSMTP)
 Server.DireccionEMAILAdministrador = Trim(Configuracion.DireccionEMAILAdministrador)
 
End Sub
Sub LimpiarUsuario(Usuario As Integer)

  ' **************************************************************
  ' Borra los Datos del Usuario, ubicado en el Array de Usuario
  ' **************************************************************
  With Usuarios(Usuario)
   .DireccionDeEmail = ""
   .FechaDeNacimiento = ""
   .Edad = ""
   .EstadoActualNumero = 0
   .EstadoActualTexto = ""
   .EstadoCivil = ""
   .Humor = ""
   .IDAliasUsuario = ""
   .IDNumericoUsuario = 0
   .Password = ""
   .Intencion = ""
   .Ocupacion = ""
   .OtraInfo = ""
   .PortActual = 0
   .Sexo = ""
   .SigNo = ""
   .Telefono = ""
   .UbicacionGeografica = ""
   .ListadoDeAmigos = ""
   .MensajesOffline = ""
   .UsuarioBloqueado = False
   .ApellidoYNombre = ""
  End With
  ' **************************************************************
  
End Sub
Function CompletarCadena(Cadena As String, Largo As Integer, Lado As String, Caracter As String) As String
 Dim contador As Integer
 Dim CadenaFinal As String
 
  ' **************************************************************
  ' Completa una Cadena del Lado Definido ([D]erecha o [I]zquierda
  ' con el [Caracter] especificado...
  ' **************************************************************
  CadenaFinal = Cadena
  For contador = 1 To (Largo - Len(Cadena))
   If UCase(Lado) = "D" Then CadenaFinal = CadenaFinal & Caracter
   If UCase(Lado) = "I" Then CadenaFinal = Caracter & CadenaFinal
  Next
  
  ' **************************************************************
  ' Devuelve la Cadena
  ' **************************************************************
  CompletarCadena = CadenaFinal
 
End Function
Function BuscarUserIDEnUsuarios(UserID As Integer) As Integer
' Dim Contador, Posicion As Integer
 
 
 ' **************************************************************
 ' Busca la Posicion Relativa de un UsuarioID dentro de la Matriz
 ' de Usuarios del Sistema
 ' **************************************************************
 'Posicion = 0
 'For Contador = 1 To Configuracion.CantidadDeUsuarios
 ' If Usuarios(Contador).IDNumericoUsuario = UserID Then
 '  Posicion = Contador
 '  Exit For
 ' End If
 'Next
 
 ' Devuelve la Posicion dentro de la Matriz
 'BuscarUserIDEnUsuarios = Posicion
 
 BuscarUserIDEnUsuarios = UserID
 
End Function
Function BuscarSocketDisponible() As Integer
Dim contador, SocketNro As Integer

 ' **************************************************************
 ' Busca el Primer Socket Disponible para derivar la Coneccion
 ' **************************************************************
 SocketNro = 0
 For contador = 1 To Configuracion.UsuariosSoportados
  If Sockets(contador).EstadoDelPort = 0 Then
   SocketNro = contador
   Exit For
  End If
 Next
 
 ' Devuelve el Port Disponible...
 BuscarSocketDisponible = SocketNro

 
 
End Function
Function NuloANada(Valor As Variant) As Variant

 If IsNull(Valor) Then
   NuloANada = ""
  Else
   NuloANada = Valor
 End If
 
End Function

Function BuscarUsuarioAliasEnUsuarios(ByRef AliasUsuario As String) As Integer
Dim contador, Posicion As Integer
 
 ' **************************************************************
 ' Descarta los usuarios Nulos ya que en la base, los
 ' usuarios disponibles (A nivel Base) aparecen con Usuario Nulo...
 ' Ya que la base contiene los N usuarios, detectando los
 ' disponibles con usuario nulo
 ' **************************************************************
 If Trim(AliasUsuario) = "" Then
  BuscarUsuarioAliasEnUsuarios = 0
  Exit Function
 End If
 
 ' **************************************************************
 ' Busca la Posicion Relativa de un AliasUsuario dentro de la Matriz
 ' de Usuarios del Sistema
 ' **************************************************************
 Posicion = 0
 For contador = 1 To Configuracion.CantidadDeUsuarios
  If Trim(UCase(Usuarios(contador).IDAliasUsuario)) = Trim(UCase(AliasUsuario)) Then
   Posicion = contador
   Exit For
  End If
 Next
 
 ' Devuelve la Posicion dentro de la Matriz
 BuscarUsuarioAliasEnUsuarios = Posicion
 
End Function
Sub CargarListadoUsuariosActuales()
Dim contador, Cantidad As Integer
Dim Texto As String

 ' **************************************************************
 ' Este sub carga los usuarios actuales logueados al sistema...
 ' (Es decir Validados)...
 ' **************************************************************
 ' Elimina el Actual Listado
 Server.ListadoUsuariosActuales.Clear
 
 ' **************************************************************
 ' Si el Sistema esta Abajo, solo muestra un Cartel de Off-Line
 ' **************************************************************
 If Configuracion.EstadoDelSistema = "Down" Then
  Server.ListadoUsuariosActuales.AddItem ("System Off-Line...")
  Exit Sub
 End If
 
 ' Carga los Usuarios Logueados
 Cantidad = 0
 For contador = 1 To Configuracion.UsuariosSoportados
  If Sockets(contador).EstadoDelPort = 2 Then
   Texto = "- " & CompletarCadena(Sockets(contador).IDAliasUsuario, 16, "D", " ") & " En Socket: [" & CompletarCadena(CStr(contador), 5, "I", 0) & "]..." & _
           " (State : "
   ' Pone como Datos el Estado del Usuario
   Select Case Usuarios(Sockets(contador).IDNumericoUsuario).EstadoActualNumero
    Case 1:
     Texto = Texto & "Available (Normal)...)"
    Case 2:
     Texto = Texto & "Not Available...)"
    Case 3:
     Texto = Texto & "Custom - " & Usuarios(Sockets(contador).IDNumericoUsuario).EstadoActualTexto & ")"
   End Select
   Server.ListadoUsuariosActuales.AddItem (Texto)
   Cantidad = Cantidad + 1
  End If
 Next
 
 ' Si no hay usuarios logueado entonces deja un aviso...
 If Cantidad = 0 Then
  Server.ListadoUsuariosActuales.AddItem ("There Not Logued User's to The System...")
 End If
 
End Sub
Public Function EnviarPasswordAUsuario(IDUsuarioAlias As String, Usuariopassword As String, EMail As String, NombreApellido As String) As Boolean
Dim Respuesta, DireccionEmail, direccionemailnombre, ObjeTo, Emailadministrador, ServidorsmTP, Mensaje As String

 DireccionEmail = EMail
 direccionemailnombre = NombreApellido
 ObjeTo = "Password Remember - EIM (Electronic Instant Messanger)..."
 Emailadministrador = Trim(Configuracion.DireccionEMAILAdministrador)
 ServidorsmTP = Trim(Configuracion.DireccionIPSMTP)
 
 ' Prepara el Mensaje
 Mensaje = "Your EIM Alias is      : " & IDUsuarioAlias & vbCrLf & _
           "The Password is        : " & Usuariopassword & vbCrLf
 
 ' Envia el Mensaje...
 Respuesta = SMTPMail.EnviarSMTPMail(CStr(ServidorsmTP), _
                                     "EIM (Electronic Instant Messanger)", _
                                     CStr(Emailadministrador), _
                                     CStr(direccionemailnombre), _
                                     CStr(DireccionEmail), _
                                     CStr(ObjeTo), _
                                     CStr(Mensaje), _
                                     "", _
                                     "Alta", _
                                     3)
 
 
 If Mid$(Respuesta, 1, 2) = "00" Then
   EnviarPasswordAUsuario = True
  Else
   EnviarPasswordAUsuario = False
 End If
  
End Function
Public Function Mensaje_Bienvenida()
Dim MensajeBienvenida

 MensajeBienvenida = MensajeBienvenida & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}} "
 MensajeBienvenida = MensajeBienvenida & "\viewkind4\uc1\pard\fs16  {\pict\wmetafile8\picw502\pich502\picwgoal285\pichgoal285"
 MensajeBienvenida = MensajeBienvenida & "0100090000037202000000005c0200000000050000000b0200000000050000000c02f601f6015c"
 MensajeBienvenida = MensajeBienvenida & "20000430 f2000cc0000001300130000000000f601f60100000000280000001300000013000000"
 MensajeBienvenida = MensajeBienvenida & "010018000000000074040000c40e0000c40e00000000000000000000ffffffffffffffffffffff"
 MensajeBienvenida = MensajeBienvenida & "ffffffffffffffaadfffaabfd4559faa557faa559fd4aabfd4c0dcc0ffffffffffffffffffffff"
 MensajeBienvenida = MensajeBienvenida & "ffffffffffffff000000fffffffffffffffffff0fbfff0fbff55bfd4009fd4559fd4009fd455bf"
 MensajeBienvenida = MensajeBienvenida & "ff009fd4559fd4007faa559fd4aadfd4f0fbffffffffffffffffffff000000ffffffffffffffff"
 MensajeBienvenida = MensajeBienvenida & "ffaadfd4559fd455bfd455bfff00bfd455bfff00bfd455bfff00bfd455bfff559fd4557faaaabf"
 MensajeBienvenida = MensajeBienvenida & "d4f0fbfff0fbffffffff000000ffffffffffffaabfff55bfd400bfff00bfff009fd4559fd4007f"
 MensajeBienvenida = MensajeBienvenida & "d4559fd4007fd4559fd4007fd400bfd400bfff557faaaabfd4ffffffffffff000000ffffffffff"
 MensajeBienvenida = MensajeBienvenida & "ff00bfd455bfd455bfff559fd4559fd4003faa003faa553faa003faa003faa559fd4007fd455bf"
 MensajeBienvenida = MensajeBienvenida & "d400bfff557fd4aadfd4ffffff000000f0fbff55bfff55dfff00dfff55dfff009fd4003faa559f"
 MensajeBienvenida = MensajeBienvenida & "d455bfd400bfd455bfd4559fd4003faa559fd4009fff00bfd4559fd4559faaffffff000000f0fb"
 MensajeBienvenida = MensajeBienvenida & "ff55bfd455dfff55dfff55dfff003faa559fd400bfd455dfff55dfff00bfff009fd4009fd4553f"
 MensajeBienvenida = MensajeBienvenida & "aa00bfd455bfff009fd4557fd4aadfd4000000aadfff00bfff55dfff55dfff55dfff55dfff55df"
 MensajeBienvenida = MensajeBienvenida & "ff55dfff00dfff55dfff55dfff55dfff55bfff00bfff55bfd400bfff55bfd4009fd4aabfd40000"
 MensajeBienvenida = MensajeBienvenida & "0055dfd455dfff55dfff55dfff55dfff55dfff55dfff55dfff55dfff55dfff55dfff55dfff00df"
 MensajeBienvenida = MensajeBienvenida & "ff55dfff55bfff00bfd400bfff559fd4559fd400000055bfff55dfff55dfff55dfff55bfd4007f"
 MensajeBienvenida = MensajeBienvenida & "aa557faa55bfd455dfff55dfff55dfff00bfd4557faa555faa009fd455bfff55bfff00bfd4557f"
 MensajeBienvenida = MensajeBienvenida & "aa00000055bfd455dfff55dfff55ffff555f7f001f7f001f55007faa55dfff55dfff55dfff557f"
 MensajeBienvenida = MensajeBienvenida & "7f001f7f001f55007faa55bfff00bfd400bfff559faa000000aadfff55dfff55ffff55dfff007f"
 MensajeBienvenida = MensajeBienvenida & "aa001f55ffffff557faa55ffff55dfff55dfff007faa001f55ffffff559faa55bfff55bfff009f"
 MensajeBienvenida = MensajeBienvenida & "d4aabfd4000000f0fbff55bfff55dfff55ffff557faaffffffffffff557faa55ffff55dfff55df"
 MensajeBienvenida = MensajeBienvenida & "ff557faaffffffffffff557faa00dfff55bfff559fd4aadfff000000ffffff55dfff55dfff55df"
 MensajeBienvenida = MensajeBienvenida & "ff559faaffffffffffff559faa55dfff55dfff55dfff559faaf0fbfff0fbff559faa55bfff00bf"
 MensajeBienvenida = MensajeBienvenida & "d455bfd4ffffff000000ffffffffffff55bfff55ffff55bfd4557faa559faa55bfd455ffff55df"
 MensajeBienvenida = MensajeBienvenida & "ff55ffff55bfd4559faa557faa55bfd400bfff559fd4f0fbffffffff000000ffffffffffffaadf"
 MensajeBienvenida = MensajeBienvenida & "ff55dfff55ffff55ffff55dfff55ffff55ffff55dfff55dfff55dfff55dfff55dfff55dfff55bf"
 MensajeBienvenida = MensajeBienvenida & "d4aadfffffffffffffff000000ffffffffffffffffffaadfff55bfff55dfff55ffff55dfff55df"
 MensajeBienvenida = MensajeBienvenida & "ff55ffff55dfff55dfff55dfff55dfff55bfd4aadfd4ffffffffffffffffff000000ffffffffff"
 MensajeBienvenida = MensajeBienvenida & "ffffffffffffffffffff55dfff55bfff55dfff55dfff55dfff55dfff55dfff00bfd455bfffffff"
 MensajeBienvenida = MensajeBienvenida & "ffffffffffffffffffffffffff000000fffffffffffffffffffffffffffffffffffff0fbffaadf"
 MensajeBienvenida = MensajeBienvenida & "ff55dfff55bfd455bfffaadffff0fbffffffffffffffffffffffffffffffffffffff0000000300"
 MensajeBienvenida = MensajeBienvenida & "0"
 MensajeBienvenida = MensajeBienvenida & "}  Welcome to your Instant Messanger..."
 MensajeBienvenida = MensajeBienvenida & "\par Please Complete your Profile in [Configuration] -> [Change My Profile...]\f1\fs17"
 MensajeBienvenida = MensajeBienvenida & "\par }"

 Mensaje_Bienvenida = MensajeBienvenida
 
End Function
