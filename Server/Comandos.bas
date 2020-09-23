Attribute VB_Name = "Comandos"
Option Explicit
' **************************************************************
' En este modulo se procesan los comandos recibidos del Cliente
' **************************************************************
'   Comando 0:      Paquetes de Login
'   Comando 1:      Cambio y Solicitud de Estado (0 - Cambio / 1 - Solicitud)
'   Comando 2:      Pide Datos usuario / Graba Datos usuario
'                   Pide Listado de Amigos / Graba Listado de Amigos
'   Comando 3:      Intercambiar Paquete  (El Servidor no Interviene)
'   Comando 4:      Mensajes
'   Comando 5:      Mensajes OnLine
' **************************************************************
Function ComandoAccion_0(Datos As Variant, Port As Integer) As String
Dim Comando As String
' **************************************************************
' Comandos
'  Comando 0: Realiza la Validacion del Usuario
'  Comando 1: Cambia la Password
'  Comando 2: Envia la Password por E-Mail...

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Datos) < 1 Then
  ' El paquete es Incorrecto
  Exit Function
 End If
 
 ' **************************************************************
 ' Separa el Comando de Loguin
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 
 Select Case Comando
  Case "0":
   Loguin_Comando0 Mid$(Datos, 2), Port
  Case "1":
   Loguin_Comando1 Mid$(Datos, 2), Port
  Case "2":
   Loguin_Comando2 Mid$(Datos, 2), Port
 End Select
 
End Function
Function ComandoAccion_1(Datos As Variant, Port As Integer) As String
Dim Comando As String
' **************************************************************
' Comandos
'  Comando 0: Cambia el Estado
'  Comando 1: Pide el Estado de Un Usuario
 
 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Datos) < 1 Then
  ' El paquete es Incorrecto
  Exit Function
 End If
 
 ' **************************************************************
 ' Saca el Comando de Estado
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 
 Select Case Comando
  Case "0":
   Estado_Comando0 Mid$(Datos, 2), Port
  Case "1":  ' ****** NO IMPLEMENTADO ******
   Estado_Comando1 Mid$(Datos, 2), Port
 End Select
 
End Function
Function ComandoAccion_2(Datos As Variant, Port As Integer) As String
Dim Comando As String
' **************************************************************
' Comandos
'  Comando 0: Pide los Datos de Un Usuario
'  Comando 1: Graba los Datos de un Usuario
'  Comando 2: Listado de Usuario - NO IMPLEMENTADO
'  Comando 3: Graba Listado de Usuario - NO IMPLEMENTADO
'  Comando 4: Verifica un Usuario Nombre
'  Comando 5: Busca Amigos

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Datos) < 1 Then
  ' El paquete es Incorrecto
  Exit Function
 End If
 
 ' **************************************************************
 ' Saca el Comando de Estado
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 
 Select Case Comando
  Case "0":
   Listado_Comando0 Mid$(Datos, 2), Port
  Case "1":
   Listado_Comando1 Mid$(Datos, 2), Port
  Case "2":
   Listado_Comando2 Mid$(Datos, 2), Port
  Case "3":
   Listado_Comando3 Mid$(Datos, 2), Port
  Case "4":
   Listado_Comando4 Mid$(Datos, 2), Port
  Case "5":
   Listado_Comando5 Mid$(Datos, 2), Port
 End Select

End Function
Function ComandoAccion_3(Datos As Variant, Port As Integer) As String
Dim Para, Mensaje, PaqueteaEnviar, De, Paquete As String
Dim PortDePara, PortdeParatemp, Respuesta As Integer
Dim tiempoinicial As Date

 ' Paquete recibido: (El Paquete Enviado es Idem)
 '      Para 16 Caracteres / DeQuen - En Paquete Enviado
 '      Paquete el Resto
 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Datos) < 17 Then
  ' El paquete es Incorrecto
  Exit Function
 End If
 Paquete = Datos
 
 ' **************************************************************
 ' Separa el Paquete
 ' **************************************************************
 Para = Trim(Mid$(Paquete, 1, 16))
 Mensaje = Trim(Mid$(Paquete, 17))
 
 ' **************************************************************
 ' Busca el Port del Usuario Destino
 ' **************************************************************
 PortdeParatemp = Varios.BuscarUsuarioAliasEnUsuarios(CStr(Para))
 If PortdeParatemp = 0 Then
  ' El Usuario no Existe - Devuelve un Paquete Informandolo
  EnviarPaqueteTCP "313" & CompletarCadena(CStr(Para), 16, "D", " "), Port
  Exit Function
 End If
 
 ' **************************************************************
 ' Verifica el Estado
 ' **************************************************************
 PortDePara = Usuarios(PortdeParatemp).PortActual
 Respuesta = 0
 If Usuarios(PortdeParatemp).EstadoActualNumero = 0 Then Respuesta = 1
 If Usuarios(PortdeParatemp).EstadoActualNumero = 2 Then Respuesta = 2
 If Usuarios(PortdeParatemp).PortActual = 0 Then Respuesta = 1
 
 ' **************************************************************
 ' Enviar el Paquete al Emisor
 ' **************************************************************
 If UCase(Trim(De)) = UCase(Trim(Para)) Then ' SI el De y Para sonI
  tiempoinicial = Time
  Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
  Loop
 End If
 EnviarPaqueteTCP "31" & Respuesta & CompletarCadena(CStr(Para), 16, "D", " "), Port
  ' 0 Ok
  ' 1 No Conectado
  ' 2 No Disponible
  ' 3 Usuario No Existe
 
 ' **************************************************************
 ' Prepara el Paquete
 ' **************************************************************
 De = Sockets(Port).IDAliasUsuario
 PaqueteaEnviar = CompletarCadena(CStr(De), 16, "D", " ") & Mensaje
 
 ' **************************************************************
 ' Espera un Segundo cuando el Emisor y el Receptor es el Mismo
 ' **************************************************************
 If UCase(Trim(De)) = UCase(Trim(Para)) Then
  tiempoinicial = Time
  Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
  Loop
 End If
 
 ' **************************************************************
 ' Enviar paquete al Receptor
 ' **************************************************************
 If PortDePara <> 0 Then
  EnviarPaqueteTCP "32" & PaqueteaEnviar, CInt(PortDePara)
 End If
 
End Function
Function ComandoAccion_4(Datos As Variant, Port As Integer) As String
Dim Para As String
'Dim Paquete, PaqueteAEnviar As String
Dim PaqueteaEnviar, Coamndo As String
Dim Paquete As Variant
Dim De As String
Dim PortdeParatemp, PortDePara As Integer
Dim Respuesta As Integer
Dim RespuestaMultichat As String

 ' Paquete enviado al Emisor:
 '      4M define que es un Mensaje
 '      0 Enviado OK
 '      1 No Conectado
 '      2 No Disponible
 '      3 Usuario No Existe
 '      M Mensaje Enviado
 '              16 Caracteres Emiso, Resto Mensaje
 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Datos) < 18 Then
  ' El paquete es Incorrecto
  Exit Function
 End If
 
 ' **************************************************************
 ' Define para Quien es el Paquete y en que port esta el Destino
 ' **************************************************************
 Dim Comando As String
 Comando = Mid$(Datos, 1, 1)
 Datos = Mid$(Datos, 2)
 Para = Trim(Mid$(Datos, 1, 16))
 Paquete = Mid$(Datos, 17)
 De = Sockets(Port).IDAliasUsuario
 PortdeParatemp = Varios.BuscarUsuarioAliasEnUsuarios(Para)
 
 ' **************************************************************
 ' Define un Evento a Enviar a Varios
 ' **************************************************************
 If Comando = "3" Then
  RespuestaMultichat = Comandos.EnviarEventoMultiUsuarios(Datos, CInt(PortdeParatemp), Port)
  ' No Contesta Nada ya que no es necesario...
  Exit Function
 End If
 
 ' **************************************************************
 ' Define el Envio de el Listado De Usuarios Multichat
 ' **************************************************************
 If Comando = "4" Then
  RespuestaMultichat = Comandos.EnviarListadoMultiUsuarios(Datos, CInt(PortdeParatemp), Port)
  ' No Contesta Nada ya que lo envio en el Modulo (Enviar.....)
  Exit Function
 End If
  
 ' **************************************************************
 ' Define Cuando es un MultiMensaje
 ' **************************************************************
 If Comando = "1" Then
  RespuestaMultichat = Comandos.EnviarMensajeMultiUsuarios(Datos, CInt(PortdeParatemp), Port)
  ' **************************************************************
  ' Espera un Segundo cuando el Emisor y el Receptor es el Mismo
  ' **************************************************************
  If UCase(Trim(De)) = UCase(Trim(Para)) Then
   Dim tiempoinicial As Date
   tiempoinicial = Time
   Do Until DateDiff("s", tiempoinicial, Time) > 1
     DoEvents
   Loop
  End If
  EnviarPaqueteTCP "4M" & Respuesta & CompletarCadena(Para, 16, "D", " ") & RespuestaMultichat, Port
  Exit Function
 End If
  
 ' **************************************************************
 ' Verificar Estado del usuario...
 ' **************************************************************
 Respuesta = Comandos.VerificaEstadoDelUsuario(CInt(PortdeParatemp))
 ' Verifica que el Usuario Existe
 'If PortdeParatemp <> 0 Then
 '  ' Usuario Existe
 '  PortDePara = Usuarios(PortdeParatemp).PortActual
 '  Respuesta = 0
 '  If Usuarios(PortdeParatemp).EstadoActualNumero = 0 Then Respuesta = 1
 '  If Usuarios(PortdeParatemp).EstadoActualNumero = 2 Then Respuesta = 2
 '  If Usuarios(PortdeParatemp).PortActual = 0 Then Respuesta = 1
 ' Else
 '  ' Usuario No Existe
 '  Respuesta = 3
 'End If
 If Respuesta <> 3 Then
  PortDePara = Usuarios(PortdeParatemp).PortActual
 End If
 
 ' **************************************************************
 ' Prepara el Paquete
 ' **************************************************************
 'Debug.Print Paquete
 PaqueteaEnviar = CompletarCadena(De, 16, "D", " ") & Paquete
 
 ' **************************************************************
 ' Enviar el Paquete al Emisor
 ' **************************************************************
 EnviarPaqueteTCP "4M" & Respuesta & CompletarCadena(Para, 16, "D", " ") & "Oki", Port
 
 ' **************************************************************
 ' Si esta Desconectado, No Disponible,o no Existe no Manda el
 ' Paquete al receptor
 ' **************************************************************
 If Respuesta = 1 Or Respuesta = 2 Or Respuesta = 3 Then Exit Function
 
 ' **************************************************************
 ' Espera un Segundo cuando el Emisor y el Receptor es el Mismo
 ' **************************************************************
 If UCase(Trim(De)) = UCase(Trim(Para)) Then
  tiempoinicial = Time
  Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
  Loop
 End If
 
 ' **************************************************************
 ' Enviar paquete al Receptor
 ' **************************************************************
 If PortDePara <> 0 Then
  EnviarPaqueteTCP "4MM" & PaqueteaEnviar, PortDePara
 End If


End Function
Function ComandoAccion_5(Datos As Variant, Port As Integer) As String
Dim Para, Paquete, De As String
Dim Contador As Integer
Dim bandera As Boolean
Dim tiempoinicial As Date
Dim Respuesta As Integer

 ' **************************************************************
 ' Largo Valido?
 ' **************************************************************
 If Len(Trim(Datos)) < 17 Then
  Exit Function ' Demasiado Corto...
 End If
 
 ' **************************************************************
 ' Define los Datos Correspondientes...
 ' **************************************************************
 Para = Trim(Mid$(Datos, 1, 16))
 Paquete = Mid$(Datos, 17)
 De = Sockets(Port).IDAliasUsuario
 
 ' **************************************************************
 ' Verifica que el usuario exista...
 ' **************************************************************
 bandera = False
 For Contador = 1 To Configuracion.CantidadDeUsuarios
  With Usuarios(Contador)
   If UCase(Trim(.IDAliasUsuario)) = UCase(Trim(Para)) Then
    bandera = True
    Exit For
   End If
  End With
 Next
 
 ' **************************************************************
 ' Si Emisor y Receptor son Iguales espera un segundo...
 ' **************************************************************
 If UCase(Trim(De)) = UCase(Trim(Para)) Then
  tiempoinicial = Time
  Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
  Loop
 End If
 
 ' **************************************************************
 ' El Usuario no Existe....
 ' **************************************************************
 If bandera = False Then ' El Usuario Para no existe...
  EnviarPaqueteTCP "32" & CompletarCadena("EIM", 16, "D", " ") & "51" & CompletarCadena(CStr(Para), 16, "D", " ") & "0", Port
  Exit Function ' SALIR !
 End If
 
 ' **************************************************************
 ' El Usuario Esta Conectado...
 ' **************************************************************
 Respuesta = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Para))
 If Usuarios(Respuesta).PortActual <> 0 Then
   ' **************************************************************
   ' El Usuario esta Conectado...
   ' **************************************************************
   Dim Usuario, FechaYHora, Mensaje As String
   Usuario = Trim(De)
   FechaYHora = Varios.FechaActualFormatoOffLine
   Mensaje = Trim(Paquete)
   EnviarPaqueteTCP "32" & CompletarCadena("EIM", 16, "D", " ") & "52" & CompletarCadena(CStr(Usuario), 16, "D", " ") & CompletarCadena(CStr(FechaYHora), 19, "D", " ") & Mensaje, Usuarios(Respuesta).PortActual
   ' **************************************************************
   ' Espera un segundo...
   ' **************************************************************
   tiempoinicial = Time
   Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
   Loop
   ' **************************************************************
   ' Obviamente como Tienen mal el Estado del Usuario, le manda
   ' el Estado...
   ' **************************************************************
   Listado_Comando2 "", Port
  Else
   ' **************************************************************
   ' El Usuario no esta Conectado...
   ' **************************************************************
   ' Graba el Nuevo Mensaje Offline...
   AgregarMensajeOffline CStr(Para), CStr(Paquete), CStr(De)
 End If
 
 ' **************************************************************
 ' Espera un segundo...
 ' **************************************************************
 tiempoinicial = Time
 Do Until DateDiff("s", tiempoinicial, Time) > 1
  DoEvents
 Loop
 
 ' **************************************************************
 ' Confirma el Correcto envio del Mensaje... (OK)
 ' **************************************************************
 EnviarPaqueteTCP "32" & CompletarCadena("EIM", 16, "D", " ") & "51" & CompletarCadena(CStr(Para), 16, "D", " ") & "1", Port
 
End Function
Function ComandoAccion_6(Datos As Variant, Port As Integer) As String


End Function
Function EnviarEventoMultiUsuarios(Datos As Variant, PortdeParatemp As Integer, Port As Integer) As String
Dim Contador, Respuesta, PortDePara As Integer
Dim CantidadDeUsuarios, RespuestaEmisor As String
Dim IDAliasUsuario, Emisor, Handle, Mensaje, De, Para As String
Dim DatosTemp, PaqueteaEnviar As String
Dim DatosAdicionales As String
Dim Evento As Integer
Dim MasDatos As String

 ' Largo Invalido
 If Len(Datos) < 21 Then Exit Function
 De = Sockets(Port).IDAliasUsuario
 CantidadDeUsuarios = Trim(Mid$(Datos, 18, 2))
 DatosTemp = Trim(Mid$(Datos, 20))
 
 ' La cantidad de Usuarios es incorrecta...
 If Not IsNumeric(CantidadDeUsuarios) Then Exit Function
 
 ' Largo Invalido
 If Len(DatosTemp) < CInt(CantidadDeUsuarios) * 27 Then Exit Function
 
 For Contador = 1 To CInt(CantidadDeUsuarios)
   Para = Mid$(DatosTemp, 1 + ((Contador - 1) * 27), 16)
   Handle = Mid$(DatosTemp, 17 + ((Contador - 1) * 27), 10)
   Evento = Mid$(DatosTemp, 27 + ((Contador - 1) * 27), 1)
   MasDatos = ""
   If Len(DatosTemp) > 27 + ((Contador - 1) * 27) Then
    MasDatos = Mid$(DatosTemp, 28 + ((Contador - 1) * 27))
   End If
   PortdeParatemp = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Para))
   Respuesta = Comandos.VerificaEstadoDelUsuario(PortdeParatemp)
   
  ' **************************************************************
  ' Define al Port donde debe enviar el Mensaje
  ' **************************************************************
  If Respuesta <> 3 Then
   PortDePara = Usuarios(PortdeParatemp).PortActual
  End If
  
  ' **************************************************************
  ' Prepara el Paquete ' 32Vane            700000000001
  ' **************************************************************
  PaqueteaEnviar = "32" & CompletarCadena(CStr(De), 16, "D", " ") & "7" & Handle & Evento
  If MasDatos <> "" Then
   PaqueteaEnviar = PaqueteaEnviar & MasDatos
  End If
   
  ' **************************************************************
  ' Si esta Desconectado, No Disponible,o no Existe no Manda el
  ' Paquete al receptor
  ' **************************************************************
  If Respuesta <> 1 And Respuesta <> 2 And Respuesta <> 3 Then
   ' **************************************************************
   ' Enviar paquete al Receptor
   ' **************************************************************
   If PortDePara <> 0 Then
    EnviarPaqueteTCP PaqueteaEnviar, PortDePara
   End If
  End If
  
  ' **************************************************************
  ' Espera un segundo...
  ' **************************************************************
  Dim tiempoinicial As Date
  tiempoinicial = Time
  Do Until DateDiff("s", tiempoinicial, Time) > 1
   DoEvents
  Loop
 
  
 Next

 
End Function
Public Function VerificaEstadoDelUsuario(PortdeParatemp As Integer) As Integer

 ' Verifica que el Usuario Existe
 If PortdeParatemp <> 0 Then
   ' Usuario Existe
   If Usuarios(PortdeParatemp).EstadoActualNumero = 0 Then VerificaEstadoDelUsuario = 1
   If Usuarios(PortdeParatemp).EstadoActualNumero = 2 Then VerificaEstadoDelUsuario = 2
   If Usuarios(PortdeParatemp).PortActual = 0 Then VerificaEstadoDelUsuario = 1
  Else
   ' Usuario No Existe
   VerificaEstadoDelUsuario = 3
 End If

End Function

Function EnviarMensajeMultiUsuarios(Datos As Variant, PortdeParatemp As Integer, Port As Integer) As String
Dim Contador, Respuesta, PortDePara As Integer
Dim CantidadDeUsuarios, RespuestaEmisor As String
Dim IDAliasUsuario, Emisor, Handle, Mensaje, De, Para As String
Dim DatosTemp, PaqueteaEnviar As String
Dim DatosAdicionales As String
Dim UltimoUsuario As String
Dim tiempoinicial As Date
    
 ' Largo Invalido
 If Len(Datos) < 21 Then Exit Function
 De = Sockets(Port).IDAliasUsuario
 CantidadDeUsuarios = Trim(Mid$(Datos, 18, 2))
 DatosTemp = Trim(Mid$(Datos, 20))
 
 ' La cantidad de Usuarios es incorrecta...
 If Not IsNumeric(CantidadDeUsuarios) Then Exit Function
 
 ' Largo Invalido
 If Len(DatosTemp) < CInt(CantidadDeUsuarios) * 26 + 2 Then Exit Function
 
 ' Define el Mensaje a Enviar....
 Mensaje = Mid$(DatosTemp, 1 + CInt(CantidadDeUsuarios) * 26)
 DatosAdicionales = Mid$(Mensaje, 1, 4)
 Mensaje = Mid$(DatosTemp, 15 + CInt(CantidadDeUsuarios) * 26)
 UltimoUsuario = ""
 
 ' Envia el Mensaje a todos y Cada uno de los Amigos...
 For Contador = 1 To CInt(CantidadDeUsuarios)
   Para = Mid$(DatosTemp, 1 + ((Contador - 1) * 26), 16)
   Handle = Mid$(DatosTemp, 17 + ((Contador - 1) * 26), 10)
   PortdeParatemp = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Para))
   Respuesta = Comandos.VerificaEstadoDelUsuario(PortdeParatemp)
   If Respuesta <> 3 Then
    PortDePara = Usuarios(PortdeParatemp).PortActual
   End If
  
  ' **************************************************************
  ' Prepara el Paquete
  ' **************************************************************
  PaqueteaEnviar = CompletarCadena(CStr(De), 16, "D", " ") & "M" & DatosAdicionales & Handle & Mensaje
 
  ' **************************************************************
  ' Si esta Desconectado, No Disponible,o no Existe no Manda el
  ' Paquete al receptor
  ' **************************************************************
  If Respuesta <> 1 And Respuesta <> 2 And Respuesta <> 3 Then
   If UCase(Trim(UltimoUsuario)) = UCase(Trim(Para)) Then
    ' Espera un Segundo...
    tiempoinicial = Time
    Do Until DateDiff("s", tiempoinicial, Time) > 1
     DoEvents
    Loop
   End If
   
   UltimoUsuario = Trim(Para)
   ' **************************************************************
   ' Enviar paquete al Receptor
   ' **************************************************************
   If PortDePara <> 0 Then
    EnviarPaqueteTCP "4MM" & PaqueteaEnviar, PortDePara
   End If
  End If
 
  ' **************************************************************
  ' Arma la Respuesta para el Emisor
  ' **************************************************************
  RespuestaEmisor = RespuestaEmisor & CompletarCadena(CStr(Respuesta), 2, "I", "0")
 
 Next

 If UCase(Trim(De)) = UCase(Trim(UltimoUsuario)) Then
  ' Espera un Segundo...
  tiempoinicial = Time
   Do Until DateDiff("s", tiempoinicial, Time) > 1
   DoEvents
  Loop
 End If
 
 ' **************************************************************
 ' Enviar el Paquete al Emisor
 ' **************************************************************
 EnviarMensajeMultiUsuarios = RespuestaEmisor
 
End Function
Function EnviarListadoMultiUsuarios(Datos As Variant, PortdeParatemp As Integer, Port As Integer) As String
Dim Contador, Respuesta, PortDePara As Integer
Dim CantidadDeUsuarios, RespuestaEmisor As String
Dim IDAliasUsuario, Emisor, Handle, Mensaje, De, Para As String
Dim DatosTemp, PaqueteaEnviar As String
Dim DatosAdicionales As String
Dim tiempoinicial As Date
Dim UltimoUsuario As String

 ' Largo Invalido
 If Len(Datos) < 30 Then Exit Function
 
 De = Sockets(Port).IDAliasUsuario
 Handle = Mid$(Trim(Datos), 16, 10)
 CantidadDeUsuarios = Trim(Mid$(Datos, 28, 2))
 Mensaje = Trim(Mid$(Datos, 30))
 
 ' La cantidad de Usuarios es incorrecta...
 If Not IsNumeric(CantidadDeUsuarios) Then Exit Function
 
 ' Largo Invalido
 If Len(Mensaje) < CInt(CantidadDeUsuarios) * 26 Then Exit Function
  
 UltimoUsuario = ""
 For Contador = 1 To CInt(CantidadDeUsuarios)
  Para = Mid$(Mensaje, 1 + ((Contador - 1) * 26), 16)
  DatosTemp = Mid$(Mensaje, 1 + ((Contador - 1) * 26) + 16, 10)
  ' **************************************************************
  ' Verifica el Estado del Usuario...
  ' **************************************************************
  PortdeParatemp = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Para))
  Respuesta = Comandos.VerificaEstadoDelUsuario(PortdeParatemp)
  ' Define el Por donde debe enviar el Mensaje...
  If Respuesta <> 3 Then
   PortDePara = Usuarios(PortdeParatemp).PortActual
  End If
  
  ' **************************************************************
  ' Prepara el Paquete
  ' **************************************************************
  PaqueteaEnviar = "32" & CompletarCadena(Trim(De), 16, "D", " ") & "4" & _
               CompletarCadena(Trim(DatosTemp), 10, "I", "0") & _
               CInt(CantidadDeUsuarios) & Mensaje
 
  ' **************************************************************
  ' Si esta Desconectado, No Disponible,o no Existe no Manda el
  ' Paquete al receptor
  ' **************************************************************
  If Respuesta <> 1 And Respuesta <> 2 And Respuesta <> 3 Then
   ' **************************************************************
   ' Enviar paquete al Receptor
   ' **************************************************************
   If PortDePara <> 0 Then
    ' **************************************************************
    ' Espera un segundo...
    ' **************************************************************
    If UCase(Trim(UltimoUsuario)) = UCase(Trim(Para)) Then
     tiempoinicial = Time
     Do Until DateDiff("s", tiempoinicial, Time) > 1
      DoEvents
     Loop
    End If
    EnviarPaqueteTCP PaqueteaEnviar, PortDePara
   End If
  End If
    
  ' **************************************************************
  ' Si el Emisor y el Recpetor son el Mismo Espera un segundo...
  ' **************************************************************
  If Trim(UCase(De)) = Trim(UCase(Para)) Then
   tiempoinicial = Time
   Do Until DateDiff("s", tiempoinicial, Time) > 1
    DoEvents
   Loop
  End If
    
  ' **************************************************************
  ' Envia la respuesta al Emisor
  ' **************************************************************
  EnviarPaqueteTCP "31" & Respuesta & CompletarCadena(CStr(Para), 16, "D", " "), Port
  ' 0 Ok
  ' 1 No Conectado
  ' 2 No Disponible
  ' 3 Usuario No Existe
      
  ' Define al Ultimo usuario que se le Mando un Paquete...
  UltimoUsuario = De
 Next
 
End Function

