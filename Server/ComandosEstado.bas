Attribute VB_Name = "ComandosEstado"
Option Explicit
Sub Estado_Comando0(Paquete As String, PortTCP As Integer)
Dim Estado, Texto As String
' Composicion de Paquete:
'                       Caracter 1       : Estado Solicitado
'                       Caracter 2 al 21 : Texto del Estado (Para Estado 3)
'
' Mensaje Enviados al Cliente:
'                       10  : Error
'                       11  : Ok
' Estados Validos:
'                   1. Visible Normal
'                   2. No Disponible
'                   3. Custom
 
 ' **************************************************************
 ' Verifica que el Paquete tenga el Largo necesario
 ' **************************************************************
 If Len(Paquete) <> 21 Then
  ' Paquete de Largo no Valido, Cancela el Proceso
  EnviarPaqueteTCP "100", PortTCP
  Exit Sub
 End If
 
 ' **************************************************************
 ' Separa el Paquete
 ' **************************************************************
 Estado = Mid$(Paquete, 1, 1)
 
 ' **************************************************************
 ' Define el Estado del Usuario
 ' **************************************************************
 Select Case Estado
  Case "1"
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualNumero = 1
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualTexto = ""
  Case "2"
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualNumero = 2
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualTexto = ""
  Case "3"
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualNumero = 3
   Texto = Trim(Mid$(Paquete, 2))
   Usuarios(Sockets(PortTCP).IDNumericoUsuario).EstadoActualTexto = Texto
  Case Else:
   EnviarPaqueteTCP "100", PortTCP
   Exit Sub
 End Select

 ' **************************************************************
 ' Envia la Confirmacion al Usuario
 ' **************************************************************
 EnviarPaqueteTCP "101", PortTCP
  
End Sub
Sub Estado_Comando1(Paquete As String, PortTCP As Integer)
Dim PaqueteEnvio As String
Dim Respuesta As Integer

 ' Devuelve:
 '          110: Usuario No Existe
 '                  16 Nombre del Usuario
 '          111: Estado del Usuario
 '                  16 NombreDelUsuario
 '                   1 Estado
 '                  20 Estado Texto
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Paquete) < 16 Then
  ' Paquete Invalido
  Exit Sub
 End If
 
 ' **************************************************************
 ' Busca el Amigo
 ' **************************************************************
 Respuesta = Varios.BuscarUsuarioAliasEnUsuarios(Trim(Mid$(Paquete, 1, 16)))
 
 ' **************************************************************
 ' Si el Usuario No Existe lo Avisa...
 ' **************************************************************
 If Respuesta = 0 Then
  PaqueteEnvio = "110" & CompletarCadena(Paquete, 16, "D", " ")
  EnviarPaqueteTCP PaqueteEnvio, PortTCP
  Exit Sub
 End If
 
 ' **************************************************************
 ' Como existe manda la Info
 ' **************************************************************
 PaqueteEnvio = "111" & _
                CompletarCadena(Usuarios(Respuesta).IDAliasUsuario, 16, "D", " ") & _
                Usuarios(Respuesta).EstadoActualNumero & _
                CompletarCadena(Usuarios(Respuesta).EstadoActualTexto, 20, "D", " ") & _
                Usuarios(Respuesta).Sexo
                
 EnviarPaqueteTCP PaqueteEnvio, PortTCP
 
End Sub

