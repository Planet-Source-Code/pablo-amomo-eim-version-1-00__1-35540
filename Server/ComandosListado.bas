Attribute VB_Name = "ComandosListado"
Option Explicit
Sub Listado_Comando5(Paquete As String, PortTCP As Integer)
Dim CadenaBuscar As String
Dim Contador, Respuesta1, Respuesta2 As Integer
Dim PaqueteaEnviar As String

' Mensaje Enviados al Cliente:
'                       250 : No Hay Coincidencias
'                       251 : Hay Coincidencias
'   Despues del 251 se envia los datos de los usuarios:
'           16 Alias Usuario
'           1  Estado
'           50 Apellido y Nombre

 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 50 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  Exit Sub
 End If
 
 ' **************************************************************
 ' Realiza la Busqueda y Devuelve el Paquete
 ' **************************************************************
 CadenaBuscar = Trim(Paquete)
 PaqueteaEnviar = ""
 For Contador = 1 To Variables.Configuracion.CantidadDeUsuarios
  Respuesta1 = InStr(1, Usuarios(Contador).ApellidoYNombre, CadenaBuscar, vbTextCompare)
  Respuesta2 = InStr(1, Usuarios(Contador).IDAliasUsuario, CadenaBuscar, vbTextCompare)
  ' Verifica que el Usuario Sea Valido... (Ya que la base contiene
  ' lo N usuarios soportados por el Sistema)
  If Trim(CStr(Usuarios(Contador).IDAliasUsuario)) <> "" Then
   If Respuesta1 <> 0 Or Respuesta2 <> 0 Then
    PaqueteaEnviar = PaqueteaEnviar & _
                     CompletarCadena(Usuarios(Contador).IDAliasUsuario, 16, "D", " ") & _
                     CompletarCadena(CStr(Usuarios(Contador).EstadoActualNumero), 1, "D", " ") & _
                     CompletarCadena(Usuarios(Contador).ApellidoYNombre, 50, "D", " ")
   End If
  End If
 Next
 
 ' **************************************************************
 ' Envia el Paquete
 ' **************************************************************
 If Trim(PaqueteaEnviar) = "" Then
   EnviarPaqueteTCP "250", PortTCP
  Else
   PaqueteaEnviar = "251" & PaqueteaEnviar
   EnviarPaqueteTCP PaqueteaEnviar, PortTCP
 End If
 
End Sub
Sub Listado_Comando4(Paquete As String, PortTCP As Integer)
Dim Respuesta As Integer
Dim PaqueteEnviar As String

' Mensaje Enviados al Cliente:
'                       240 : El Usuario No Existe
'                       241 : El Usuario Existe
'                       Ambos (240,241) se le agrega
'                       Estado:         1 Caracter
'                       Textos Estado:  20 Caracteres
 
 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 16 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  EnviarPaqueteTCP "240", PortTCP
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica si existe el Usuario
 ' **************************************************************
 Respuesta = CInt(Varios.BuscarUsuarioAliasEnUsuarios(Trim(Paquete)))
 
 ' **************************************************************
 ' Envia la respuesta
 ' **************************************************************
 If Respuesta <> 0 Then
   PaqueteEnviar = "241" & CompletarCadena(CStr(Usuarios(Respuesta).EstadoActualNumero), 1, "D", " ") & _
                   CompletarCadena(Usuarios(Respuesta).EstadoActualTexto, 20, "D", " ") & _
                   Usuarios(Respuesta).Sexo & _
                   CompletarCadena(CStr(Usuarios(Respuesta).IDAliasUsuario), 16, "D", " ")
                   
   EnviarPaqueteTCP PaqueteEnviar, PortTCP
  Else
   PaqueteEnviar = "240" & "0" & CompletarCadena("", 20, "D", " ")
   EnviarPaqueteTCP PaqueteEnviar, PortTCP
 End If
 
End Sub
Sub Listado_Comando2(Paquete As String, PortTCP As Integer)
Dim Listado, Amigos() As String
Dim Contador, CantidadAmigosGrupos, PosicionCadena, PosicionUsuario As Integer
Dim Usuario As String

' Formato del Pauqte
'   NombreDelGrupo 20
'   IDNombreDelAmigo 16
'   EstadoDelAmigoEstado 1
'   EstadoDelAmigoTexto 20
'   NombreDelAmigo 50
'   Sexo 1
'   Existe 1

 ' **************************************************************
 ' Carga el Listado actual de los Amigos del Usuario
 ' **************************************************************
 Listado = Usuarios(Sockets(PortTCP).IDNumericoUsuario).ListadoDeAmigos
 
 ' **************************************************************
 ' Verifica que tenga algo que enviar
 ' **************************************************************
 ' Esto se saco para cuando el Listado es nulo (Es decir, No tiene
 ' Amigos...
 'If Len(Listado) = 0 Then
 ' Exit Sub
 'End If
 
 ' **************************************************************
 ' La Cantidad de Amigos/Grupos
 ' **************************************************************
 CantidadAmigosGrupos = 0
 For Contador = 1 To Len(Usuarios(Sockets(PortTCP).IDNumericoUsuario).ListadoDeAmigos)
  If Mid$(Listado, Contador, 1) = ";" Then
   CantidadAmigosGrupos = CantidadAmigosGrupos + 1
  End If
 Next
   
 ' **************************************************************
 ' Separa los Usuario para empezar a Procesar
 ' **************************************************************
 Amigos = Split(Listado, ";")
  
 Listado = ""
 For Contador = 1 To CantidadAmigosGrupos
  PosicionCadena = InStr(Amigos(Contador - 1), "@")
  PosicionUsuario = 0
  If PosicionCadena <> 1 And PosicionCadena <> 0 Then
   PosicionUsuario = Varios.BuscarUsuarioAliasEnUsuarios(Mid$(Amigos(Contador - 1), 1, PosicionCadena - 1))
  End If
  ' **************************************************************
  ' Arma el Usuario a Enviar on los datos si tiene usuario
  ' Primero pone el Usuario y Grupo
  ' **************************************************************
  ' Agrega el Grupo
  If Len(Amigos(Contador - 1)) = PosicionCadena Then
    Listado = Listado & CompletarCadena(" ", 20, "D", " ")
   Else
    Listado = Listado & CompletarCadena(Mid$(Amigos(Contador - 1), PosicionCadena + 1), 20, "D", " ")
  End If
  ' Agrega el UsuarioIDAlias
  If PosicionCadena <> 1 And PosicionCadena <> 0 Then
    Listado = Listado & CompletarCadena(Mid$(Amigos(Contador - 1), 1, PosicionCadena - 1), 16, "D", " ")
   Else
    Listado = Listado & CompletarCadena(" ", 16, "D", " ")
  End If
  ' **************************************************************
  ' Agrega los Datos Genericos si el Usuario exsite...
  ' **************************************************************
  If PosicionUsuario <> 0 Then
   ' Agrega Estado del Amigo
   Listado = Listado & CompletarCadena(CStr(Usuarios(PosicionUsuario).EstadoActualNumero), 1, "D", " ")
   ' Agrega EstadoTexto del Amigo
   Listado = Listado & CompletarCadena(Usuarios(PosicionUsuario).EstadoActualTexto, 20, "D", " ")
   ' Agrega Nombre del Amigo
   Listado = Listado & CompletarCadena(Usuarios(PosicionUsuario).ApellidoYNombre, 50, "D", " ")
   ' Agrega El Sexo
   Listado = Listado & CompletarCadena(Usuarios(PosicionUsuario).Sexo, 1, "D", " ")
   ' Agrega El Usuario Existe 1=si 0=no
   Listado = Listado & "1"
   ' Agrega la Direccion de EMail
   Listado = Listado & CompletarCadena(Usuarios(PosicionUsuario).DireccionDeEmail, 50, "D", " ")
  End If
  ' **************************************************************
  ' Agrega los Datos Genericos si el Usuario no Existe o es
  ' un grupo...
  ' **************************************************************
  If PosicionUsuario = 0 Then
   ' Agrega Estado del Amigo
   Listado = Listado & CompletarCadena("0", 1, "D", " ")
   ' Agrega EstadoTexto del Amigo
   Listado = Listado & CompletarCadena(" ", 20, "D", " ")
   ' Agrega Nombre del Amigo
   Listado = Listado & CompletarCadena(" ", 50, "D", " ")
   ' Agrega El Sexo
   Listado = Listado & CompletarCadena("M", 1, "D", " ")
   ' Agrega El Usuario Existe 1=si 0=no
   Listado = Listado & "0"
   ' Agrega la Direccion de EMail
   Listado = Listado & CompletarCadena(" ", 50, "D", " ")
  End If
 Next

 ' **************************************************************
 ' Envia el Paquete
 ' **************************************************************
 ' Listado = "22" & Usuarios(Sockets(PortTCP).IDNumericoUsuario).ListadoDeAmigos
 Listado = "22" & Listado
 EnviarPaqueteTCP Listado, PortTCP
   
End Sub
Sub Listado_Comando3(Paquete As String, PortTCP As Integer)
Dim Respuesta As Integer

' Mensaje Enviados al Cliente:
'                       230 : Error en Grabacion
'                       231 : Grabacion Ok

 
 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' Esto no es necesario ya que si borra la Lista el Paquete es 0
 ' **************************************************************
 'If Len(Trim(Paquete)) = 0 Then
 ' ' Paquete de Largo no Valido, Cancela el Proceso
 ' EnviarPaqueteTCP "230", PortTCP
 ' Exit Sub
 'End If
 
 ' **************************************************************
 ' Pone el Nuevo Listado de Usuario
 ' **************************************************************
 Usuarios(Sockets(PortTCP).IDNumericoUsuario).ListadoDeAmigos = Trim(Paquete)
 
 ' **************************************************************
 ' Graba en la Base
 ' **************************************************************
 Respuesta = BaseDeDatos.GrabarListadoDeAmigos(Sockets(PortTCP).IDNumericoUsuario, Paquete)
 
 ' **************************************************************
 ' Confirma la Grabacion
 ' **************************************************************
 If Respuesta = 1 Then
   EnviarPaqueteTCP "231", PortTCP
  Else
   EnviarPaqueteTCP "230", PortTCP
 End If
 
End Sub
Sub Listado_Comando0(Paquete As String, PortTCP As Integer)
Dim UsuarioPosicion As Integer
Dim UsuarioNombre, PaqueteDatos As String
' Composicion de Paquete:
'                       Caracter 1 al 16    :  IDAliasUsuario
'
' Mensaje Enviados al Cliente:
'                       00  : El Usuario no existe
'                       01  : Paquete con los Datos
'                       02  : EL Usuario no Solicitado Existe

 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 16 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  EnviarPaqueteTCP "200", PortTCP
  Exit Sub
 End If

 UsuarioNombre = Paquete
 ' **************************************************************
 ' Busca la Posicion relativa del Usuario
 ' **************************************************************
 UsuarioPosicion = Varios.BuscarUsuarioAliasEnUsuarios(Trim(UsuarioNombre))
 ' No se encontro el Usuario Solicitado
 If UsuarioPosicion = 0 Then
  EnviarPaqueteTCP "202" & UsuarioNombre, PortTCP
  Exit Sub
 End If
 
 ' **************************************************************
 ' Envia el Paquete (Con los Datos solicitados)
 ' **************************************************************
 ' Prepara el Paquete con los Datos del Usuario
 PaqueteDatos = "201" & UsuarioNombre
 PaqueteDatos = PaqueteDatos & CompletarCadena(Usuarios(UsuarioPosicion).ApellidoYNombre, 50, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).DireccionDeEmail, 50, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Edad, 2, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Sexo, 1, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).UbicacionGeografica, 20, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Intencion, 20, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Humor, 20, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Ocupacion, 20, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).SigNo, 15, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).EstadoCivil, 1, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).Telefono, 50, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).OtraInfo, 150, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).FechaDeNacimiento, 10, "D", " ") & _
                CompletarCadena(CInt(Usuarios(UsuarioPosicion).EstadoActualNumero), 1, "D", " ") & _
                CompletarCadena(Usuarios(UsuarioPosicion).EstadoActualTexto, 20, "D", " ")
 
 ' Envia el Paquete
 EnviarPaqueteTCP PaqueteDatos, PortTCP
 
End Sub
Sub Listado_Comando1(Paquete As String, PortTCP As Integer)
Dim UsuarioPosicion, Respuesta As Integer

' Composicion de Paquete:
'                       Todos los Campos...
'
' Mensaje Enviados al Cliente:
'                       1  : Grabado OK...
'                       0  : Error No se pudo Grabar...

 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 409 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  Exit Sub
 End If

 ' **************************************************************
 ' Guarda en Usuarios los Nuevos Datos del Mismo...
 ' **************************************************************
 UsuarioPosicion = Sockets(PortTCP).IDNumericoUsuario
 With Usuarios(UsuarioPosicion)
  .ApellidoYNombre = Trim(Mid$(Paquete, 1, 50))
  .DireccionDeEmail = Trim(Mid$(Paquete, 51, 50))
  .Edad = Trim(Mid$(Paquete, 101, 2))
  .Sexo = Trim(Mid$(Paquete, 103, 1))
  .UbicacionGeografica = Trim(Mid$(Paquete, 104, 20))
  .Intencion = Trim(Mid$(Paquete, 124, 20))
  .Humor = Trim(Mid$(Paquete, 144, 20))
  .Ocupacion = Trim(Mid$(Paquete, 164, 20))
  .SigNo = Trim(Mid$(Paquete, 184, 15))
  .EstadoCivil = Trim(Mid$(Paquete, 199, 1))
  .Telefono = Trim(Mid$(Paquete, 200, 50))
  .OtraInfo = Trim(Mid$(Paquete, 250, 150))
  .FechaDeNacimiento = Trim(Mid$(Paquete, 400, 10))
 End With
 
 ' **************************************************************
 ' Graba los Datos modificados del Usuario
 ' **************************************************************
 Respuesta = BaseDeDatos.GrabarModificacionesUsuario(CInt(UsuarioPosicion))
 
 ' **************************************************************
 ' Confirma que la Actualizacion fue exitosa
 ' **************************************************************
 If Respuesta = 1 Then
   EnviarPaqueteTCP "211", PortTCP
   ' EscribirEvento "Los Datos del Usuario [" & Trim(Sockets(PortTCP).IDAliasUsuario) & "] fueron Cambiados Exitosamente...", vbBlue
   EscribirEvento "The Profile Of The User [" & Trim(Sockets(PortTCP).IDAliasUsuario) & "] was Susecfully Updated...", vbBlue
  Else
   EnviarPaqueteTCP "210", PortTCP
 End If
 
End Sub


