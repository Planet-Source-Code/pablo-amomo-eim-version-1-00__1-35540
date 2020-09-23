Attribute VB_Name = "SocketTCP"
Option Explicit
Sub InicializarSocket()

 ' **************************************************************
 ' Abre y Prepara el Primer Socket
 ' **************************************************************
 Server.TCPSocket(0).LocalPort = Configuracion.PortTCP
 Server.TCPSocket(0).Protocol = sckTCPProtocol
 Server.TCPSocket(0).Listen
  
End Sub
Function RecibirPaqueteTCP(ByRef PaqueteRecibido As Variant) As Variant

 ' **************************************************************
 ' Modulo que preprosesa lo que se recibe por el Socket
 ' **************************************************************
 RecibirPaqueteTCP = PaqueteRecibido
 ' **************************************************************
 
End Function
Function EnviarPaqueteTCP(ByRef PaqueteaEnviar As Variant, SocketNro As Integer)  ', Optional SinEspera As String) As String
On Error GoTo ErrorEnviarPaquete
'Dim TiempoInicial As Date
 
 'If UCase(SinEspera) <> "SI" Then
 ' ' **************************************************************
 ' ' Evita Solapamientos...
 ' ' **************************************************************
 ' If UltimoPort = SocketNro Then
 '  TiempoInicial = Time
 '  Do Until DateDiff("s", TiempoInicial, Time) > 1
 '   DoEvents
 '  Loop
 ' End If
 'End If
 'UltimoPort = SocketNro
 'UltimoEnvio = Time
 
 ' **************************************************************
 ' Modulo que preprosesa lo que se desea Enviar
 ' **************************************************************
 Server.TCPSocket(SocketNro).SendData (PaqueteaEnviar)
 ' **************************************************************
 
SalirEnviarPaquete:
  Exit Function
  
ErrorEnviarPaquete:
  Resume SalirEnviarPaquete
 
End Function
