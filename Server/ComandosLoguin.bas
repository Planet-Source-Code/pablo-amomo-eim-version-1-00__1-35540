Attribute VB_Name = "ComandosLoguin"
Option Explicit
Sub Loguin_Comando0(Paquete As String, PortTCP As Integer)
Dim AliasUsuario, Password, Estado, EstadoTexto As String
Dim Posicion As Integer
' Composicion de Paquete:
'                       Caracter 1 al 16    :  IDAliasUsuario
'                       Caracter 17 al 28   :  Password
'
' Mensaje Enviados al Cliente:
'                       01  : El Usuario no existe
'                       02  : Password Incorrecta
'                       03  : Loguin Correcto
'                       04  : Usuario Lockeado

 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 29 Then
  If Mid$(Paquete, 29, 1) = "3" Then
    If Len(Paquete) <> 49 Then
     Exit Sub
    End If
   Else
  ' Paquete de Largo no Valido, Cancela el Proceso
    Exit Sub
  End If
 End If
 
 ' **************************************************************
 ' Abre el Paquete
 ' **************************************************************
 AliasUsuario = Trim(CStr(Mid$(Paquete, 1, 16)))
 Password = Trim(CStr(Mid$(Paquete, 17, 12)))
 Estado = CStr(Mid$(Paquete, 29, 1))
 If Estado = "3" Then
   EstadoTexto = Trim(CStr(Mid$(Paquete, 30, 20)))
  Else
   EstadoTexto = ""
 End If
 
 ' **************************************************************
 ' Busca los datos del usuario
 ' **************************************************************
 Posicion = BuscarUsuarioAliasEnUsuarios(CStr(AliasUsuario))
 ' Verifica que el Usuario Exista
 If Posicion = 0 Then
  ' Devuelve un Mensaje informando que el usuario no Existe
  EnviarPaqueteTCP "001", PortTCP
  ' EscribirEvento "El Usuario [" & AliasUsuario & "] Ingresado Desde [" & Server.TCPSocket(PortTCP).RemoteHostIP & "] no Existe...", vbRed
  EscribirEvento "The User [" & AliasUsuario & "] Entering From [" & Server.TCPSocket(PortTCP).RemoteHostIP & "] not Exist...", vbRed
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica que el usuario no este Lockeado
 ' **************************************************************
 If Usuarios(Posicion).UsuarioBloqueado Then
  ' Devuelve un Mensaje Informando que la Password es Incorrecta
  EnviarPaqueteTCP "004", PortTCP
  ' EscribirEvento "El Usuario [" & AliasUsuario & "] se encuentra Lockeado...", vbRed
  EscribirEvento "The User [" & AliasUsuario & "] are Locked-Out...", vbRed
  Exit Sub
 End If
 
 
 ' **************************************************************
 ' Valida la Password del Usuario
 ' **************************************************************
 If Trim(Usuarios(Posicion).Password) <> Trim(DesEncriptar(CStr(Password), Usuarios(Posicion).Password)) Then
  ' Devuelve un Mensaje Informando que la Password es Incorrecta
  EnviarPaqueteTCP "002", PortTCP
  EscribirEvento "The Password of The User [" & AliasUsuario & "] From [" & Server.TCPSocket(PortTCP).RemoteHostIP & "] was Incorrect...", vbRed
  Exit Sub
 End If
 
 ' **************************************************************
 ' Valida que el Estado Sea Numerico sino descarta el Paquete
 ' **************************************************************
 If Not IsNumeric(Estado) Then
  Exit Sub
 End If
 
 Dim Sexo, NombreyApellido As String
 Sexo = Usuarios(Posicion).Sexo
 NombreyApellido = CompletarCadena(Usuarios(Posicion).ApellidoYNombre, 50, "D", " ")
 ' Loguin Correcto
 EnviarPaqueteTCP "003" & Sexo & NombreyApellido, PortTCP
 '
 ' Cargar los Datos del Usuario
 '
 Sockets(PortTCP).EstadoDelPort = 2
 Sockets(PortTCP).IDAliasUsuario = Usuarios(Posicion).IDAliasUsuario
 Sockets(PortTCP).IDNumericoUsuario = Usuarios(Posicion).IDNumericoUsuario
 Usuarios(Posicion).PortActual = PortTCP
 Usuarios(Posicion).EstadoActualNumero = CInt(Estado)
 Usuarios(Posicion).EstadoActualTexto = EstadoTexto
 EscribirEvento "The User [" & AliasUsuario & "] was Succecfully Granted from [" & Server.TCPSocket(PortTCP).RemoteHostIP & "]...", vbBlue
 ' Graba la Fecha y Hora del Logueo...
 BaseDeDatos.GrabarFechaUltimoLogueo (Posicion)
 
  
 ' **************************************************************
 ' Verifica los Mensajes Off-Line...
 ' **************************************************************
 EnviaMensajesOffline Sockets(PortTCP).IDAliasUsuario, PortTCP

  
End Sub
Sub Loguin_Comando1(Paquete As String, PortTCP As Integer)
Dim PasswordActual, PasswordNueva, PWDTemp1, PWDTemp2 As String
Dim Posicion, Respuesta As Integer
' Composicion de Paquete:
'                       Caracter 1 al 12    : Password Nueva (Encriptada con Password Actual)
'                       Caracter 13 al 24   : Password Actual (Encriptada con Ella Misma)
' Mensaje Enviados al Cliente:
'                       01  : Ok

 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 24 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  Exit Sub
 End If
 
 ' **************************************************************
 ' Abre el Paquete
 ' **************************************************************
 PasswordNueva = Trim(CStr(Mid$(Paquete, 1, 12)))
 PasswordActual = Trim(CStr(Mid$(Paquete, 13)))
 
 
 ' **************************************************************
 ' Busca los datos del usuario
 ' **************************************************************
 Posicion = BuscarUsuarioAliasEnUsuarios(CStr(Sockets(PortTCP).IDAliasUsuario))
 ' Verifica que el Usuario Exista
 If Posicion = 0 Then
  ' Si no se encuentra el Usuario se Desacarta el Paquete
  EscribirEvento "The User [" & Trim(Sockets(PortTCP).IDAliasUsuario) & "] tray to Change His Password but That User Don't Exist...", vbRed
  Exit Sub
 End If
 
 ' **************************************************************
 ' Desencripta la Password
 ' **************************************************************
 PWDTemp1 = DesEncriptar(Trim(CStr(PasswordNueva)), Trim(CStr(Usuarios(Posicion).Password)))
 PWDTemp2 = DesEncriptar(Trim(CStr(PasswordActual)), Trim(CStr(Usuarios(Posicion).Password)))

 ' **************************************************************
 ' Valida la Password Actual
 ' **************************************************************
 If Trim(Usuarios(Posicion).Password) <> Trim(PWDTemp2) Then
  ' La Password Actual no es Correcta
   EscribirEvento "The Password for The User [" & Sockets(PortTCP).IDAliasUsuario & "] Enter From [" & Server.TCPSocket(PortTCP).RemoteHostIP & "] for 'Password Change' was Not Correct...", vbRed
  Exit Sub
 End If
 
 ' **************************************************************
 ' Ejecuta el Cambio de PWD
 ' **************************************************************
 Respuesta = GrabarNuevaPassword(CInt(Sockets(PortTCP).IDNumericoUsuario), CStr(Trim(PWDTemp1)))
 
 ' Verifica si se pudo grabar la Password OK 1 - OK / 0 - Error
 If Respuesta = 0 Then
  ' Hubo un Error
  Exit Sub
 End If
 
 ' Confirma el Cambio de PWD
 EnviarPaqueteTCP "011", PortTCP
 
 ' Escribe un Evento
 EscribirEvento "The Password for The User [" & Trim(Sockets(PortTCP).IDAliasUsuario) & "] was Succecfully Changed...", vbBlue
 
End Sub
Sub Loguin_Comando2(Paquete As String, PortTCP As Integer)
Dim UsuarioPosicion As Integer
Dim ServidorsmTP, DireccionEmail, direccionemailnombre, Mensaje, Respuesta, ObjeTo, Emailadministrador As String
 
 ' Composicion del Paquete:
 '          1 a 16:     Nombre de Usuario
 ' Mensaje enviado al Cliente:
 '       0: No se pudo enviar la Password...
 '       1: La password fue enviada a enviopassworddireccionmail
 '       2: El usuario no posee direccion de Email
 '       3: EL usuario no existe...
 
 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 16 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  Exit Sub
 End If
 
 UsuarioPosicion = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Paquete))
 DireccionEmail = Trim(CStr(Usuarios(UsuarioPosicion).DireccionDeEmail))
 direccionemailnombre = Trim(CStr(Usuarios(UsuarioPosicion).ApellidoYNombre))
 ObjeTo = "Recuerdo de Password - EIM (Electronic Instant Messanger)..."
 Emailadministrador = Trim(Configuracion.DireccionEMAILAdministrador)
 ServidorsmTP = Trim(Configuracion.DireccionIPSMTP)
 
 ' **************************************************************
 ' Envia la Respuesta al Usuario
 ' **************************************************************
 ' El Usuario no posee direccion de E-Mail
 If DireccionEmail = "" Or IsNull(DireccionEmail) Then
  EnviarPaqueteTCP "022", PortTCP
  Exit Sub
 End If
 ' El Usuario no Existe...
 If UsuarioPosicion = 0 Then
  EnviarPaqueteTCP "023", PortTCP
  Exit Sub
 End If
 
 ' Prepara el Mensaje
 Mensaje = "Your EIM Alias is      : " & Trim(Usuarios(UsuarioPosicion).IDAliasUsuario) & vbCrLf & _
           "The Password is        : " & Trim(Usuarios(UsuarioPosicion).Password) & vbCrLf
 
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
   ' El mensaje se envio OK
   EnviarPaqueteTCP "021" & CompletarCadena(CStr(DireccionEmail), 50, "D", " "), PortTCP
   EscribirEvento "Was Send to [" & DireccionEmail & "] The Password of The User [" & direccionemailnombre & "]...", vbBlue
  Else
   ' No se pudo enviar el Mensaje
   EnviarPaqueteTCP "020" & CompletarCadena(CStr(DireccionEmail), 50, "D", " "), PortTCP
   EscribirEvento "Was Not Posibble Send to [" & DireccionEmail & "] the Password of The User [" & direccionemailnombre & "]...", vbRed
 End If
  
End Sub
