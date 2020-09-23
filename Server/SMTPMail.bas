Attribute VB_Name = "SMTPMail"
Option Explicit
Rem ************************************************
Rem Utilizado para verificar la respuesta desde
Rem el SMTP Server
Rem ************************************************
Public RespuestaWinsockSMTP As String
Private EventoOK As Integer
Private RespuestaEventoError As String
Private ErrorCodArc As Integer
Private Version As String
Public Function EnviarSMTPMail(EmailServidor As String, EmailDesdeNombre As String, EmailDesdeDireccion As String, EmailParaNombre As String, EMailParaDireccion As String, EmailObjetoDelMensaje As String, EmailMensaje As String, EMailArchivoAtach As String, PrioridadDeEntrega As String, EmailTimeOut As Integer) As String

 ' Respuestas de la Funcion:
 ' "00: Todo OK!!!..."
 ' "01: No fue Posible Entablar Conección Contra el Servidor [Servidor]..."
 ' "02: No fue Posible Enviar el Comando de Inicialización al Servidor SMTP..."
 ' "03: No fue Posible Enviar el Comando [MAIL FROM:] al Servidor SMTP..."
 ' "04: No fue Posible enviar el Comando [RCPT] al Servidor SMTP..."
 ' "05: No fue Posible enviar el comando [DATA] al Servidor SMTP..."
 ' "06: No fue Posible Enviar el Texto del Mensaje al Servidor SMTP..."
 ' "07: No fue Posible Realizar la Desconección del Servidor SMTP..."
 ' "08: No fue posible Abrir el Socket para el Envio..."
 ' "09: Error de Archivo..."
 ' "10: Se produjo el Error: " & ErrorRecibido
 ' "11: El Socket se Encuentra Abierto y no Puede ser Cerrado..."
 
 If Trim(EmailParaNombre) = "" Then
  EmailParaNombre = EMailParaDireccion
 End If
 
 If Trim(EmailDesdeNombre) = "" Then
  EmailDesdeNombre = EmailDesdeDireccion
 End If
 
 ' Error: si no se ingresa la direccion de destino...
 If Trim(EMailParaDireccion) = "" Then
  EnviarSMTPMail = "10: Debe Ingresar una Direccion Destino de Email..."
  Exit Function
 End If
 
 EnviarMail EmailServidor, EmailDesdeNombre, EmailDesdeDireccion, EmailParaNombre, EMailParaDireccion, EmailObjetoDelMensaje, EmailMensaje, EMailArchivoAtach, PrioridadDeEntrega, EmailTimeOut
 EnviarSMTPMail = RespuestaEventoError & "[" & EmailParaNombre & "]"
  

End Function
Private Sub EnviarMail(EmailServidor As String, EmailDesdeNombre As String, EmailDesdeDireccion As String, EmailParaNombre As String, EMailParaDireccion As String, EmailObjetoDelMensaje As String, EmailMensaje As String, EMailArchivoAtach As String, EMailPrioridad As String, EmailTimeOut As Integer)
On Error GoTo ErrorEnviando
Dim Fecha, EncabezadoMail, Mensaje, MailDe As String
Dim Casilla, Rsp, NombreArchivo As String
Dim ErrorRecibido As String
Dim A, B As Long
' Para Multiples Atach
Dim NombreArchivoAtach(100) As String
Dim CantidadArchivoAtach, Ultimo As Integer
Dim StrArchivosAtach, ArchivoTMP As String
Dim Prioridad As Integer
Dim TextoPrioridad As String

 ' ***************************************************
 ' Buscar y Definir la Prioridad
 ' ***************************************************
 Prioridad = 0
 EMailPrioridad = Trim(EMailPrioridad)
 If UCase(EMailPrioridad) = UCase("Alta") Then
  Prioridad = 1
  TextoPrioridad = "High"
 End If
 If UCase(EMailPrioridad) = UCase("Normal") Then
  Prioridad = 3
  TextoPrioridad = "Normal"
 End If
 If UCase(EMailPrioridad) = UCase("Baja") Then
  Prioridad = 5
  TextoPrioridad = "Low"
 End If
 If Prioridad = 0 Then
  Prioridad = 3
  TextoPrioridad = "Normal"
 End If
 ' ***************************************************
 
 ' ***************************************************
 ' Busca y Carga Multiples Atach
 ' ***************************************************
 StrArchivosAtach = ""
 EMailArchivoAtach = Trim(EMailArchivoAtach)
 If EMailArchivoAtach <> "" Then
  CantidadArchivoAtach = 0
  Ultimo = 0
  For A = 1 To Len(EMailArchivoAtach)
   If Mid$(EMailArchivoAtach, A, 1) = ";" Then
    If Mid$(EMailArchivoAtach, Ultimo + 1, A - (Ultimo + 1)) <> "" Then
     CantidadArchivoAtach = CantidadArchivoAtach + 1
     NombreArchivoAtach(CantidadArchivoAtach) = Mid$(EMailArchivoAtach, Ultimo + 1, A - (Ultimo + 1))
     Ultimo = A
    End If
   End If
  Next
  ' Agrega el Ultimo, y si es uno solo lo pone
  ' Solo lo agrega cuando el ultimo caracter es distinto de ;
  If Mid$(EMailArchivoAtach, Len(EMailArchivoAtach), 1) <> ";" Then
   CantidadArchivoAtach = CantidadArchivoAtach + 1
   NombreArchivoAtach(CantidadArchivoAtach) = Mid$(EMailArchivoAtach, Ultimo + 1)
  End If
  
  ' ************************************************************************
  ' Carga los Archivos Atachados
  ' ************************************************************************
  ' NOTA: Utiliza archivoTMP para traer el Error desde la codificacion del Archivo
  For A = 1 To CantidadArchivoAtach
   ArchivoTMP = CodificarArchivo(NombreArchivoAtach(A))
   If ErrorCodArc <> 0 Then
    ' Ante un Error Sale, trae la descripcion del error en ArchivoTmp
    Exit For
   End If
   StrArchivosAtach = StrArchivosAtach & ArchivoTMP & vbCrLf
  Next
  ' Si hubo un error cancela el envio
  If ErrorCodArc <> 0 Then
   RespuestaEventoError = "09: Hubo un error al Intentar Codifcar el Archivo [" & NombreArchivoAtach(A) & " - '" & ArchivoTMP & "' ]..."
   Server.SMTPMailSocket.Close
   ' ******************************
   ' Libera la Memoria
   ' ******************************
   Mensaje = ""
   ArchivoTMP = ""
   StrArchivosAtach = ""
   ' ******************************
   Exit Sub
  End If
  ArchivoTMP = ""
 End If
 ' ***************************************************
    
' ***************************************************
' Configura y Prepara el Winsock para la comunicacion
' ***************************************************
' Si esta abierto el Socket lo Cierra
 If Server.SMTPMailSocket.State <> sckClosed Then
  Server.SMTPMailSocket.Close
  ' Si el Socket esta Abierto lo cierra y espera 5 segundos
  ' a que cambie el estado a sckClosed
  EsperarEvento "241", 5
  If Server.SMTPMailSocket.State <> sckClosed Then
   RespuestaEventoError = "11: El Socket se Encuentra Abierto y no Puede ser Cerrado..."
   Mensaje = ""
   StrArchivosAtach = ""
   Exit Sub
  End If
 End If
' ***************************************************
 
 
' Define el Local Port
 Server.SMTPMailSocket.LocalPort = 0
' Define el protocolo a usar en la coneccion
 Server.SMTPMailSocket.Protocol = sckTCPProtocol
' Define el Servidor SMTP al Cual se le envia el Mensaje
 Server.SMTPMailSocket.RemoteHost = Trim(Configuracion.DireccionIPSMTP)
' Define el Port a Utilizar (Port 25 para SMTP)
 Server.SMTPMailSocket.RemotePort = 25
' ***************************************************

If Server.SMTPMailSocket.State = sckClosed Then
    
 RespuestaEventoError = "00: Todo OK!!!..."
 
 ' ***************************************************
 ' Se define la Hora y Fecha enviados en el MAIL
 ' En el "-0300" se define la diferencia Horaria
 ' ***************************************************
 Fecha = Format(Date, "Ddd") & ", " & _
         Format(Date, "dd mmm YYYY") & " " & _
         Format(Time, "hh:mm:ss") & "" & _
         " -0300"
 ' ***************************************************
 
 ' ***************************************************
 ' Prepara la info para enviar al servidor SMTP
 ' ***************************************************
 MailDe = "mail from: " & EmailDesdeDireccion & vbCrLf
 Casilla = "rcpt to: " & EMailParaDireccion & vbCrLf
 EncabezadoMail = "From: " & Chr$(34) & EmailDesdeNombre & Chr$(34) & _
                  " <" & LCase(EmailDesdeDireccion) & "> " & vbCrLf & _
                  "To: <" & EmailParaNombre & ">" & vbCrLf & _
                  "Subject: " & EmailObjetoDelMensaje & vbCrLf & _
                  "Date: " & Fecha & vbCrLf & _
                  "X-Priority: " & Prioridad & vbCrLf & _
                  "X-MSMail-Priority: " & TextoPrioridad & vbCrLf & _
                  "X-Mailer: SMTP Mail Versión " & Version & " By Pablo Amomo" & _
                  vbCrLf & vbCrLf
    
 ' ************************************************************************
 ' Prepara el Mensaje del EMAIL Con los Atach Respectivos
 ' ************************************************************************
 Mensaje = EmailMensaje
 
 ' ************************************************************************
 ' Agrega al Mensaje el o los Atach
 ' ************************************************************************
 If CantidadArchivoAtach > 0 Then
  Mensaje = Mensaje & vbCrLf & vbCrLf & vbCrLf & StrArchivosAtach
 End If
  ' Limpia la memoria del archivo atachado (O los archivos)
  StrArchivosAtach = ""
 ' ************************************************************************
    
 ' ***************************************************
 ' Se conecta contra el Servidor SMTP
 ' ***************************************************
 Server.SMTPMailSocket.Connect
 EsperarEvento "220", EmailTimeOut
 If EventoOK = 0 Then
  RespuestaEventoError = "01: No fue Posible Entablar Conección Contra el Servidor [" & EmailServidor & "]..."
  Server.SMTPMailSocket.Close
  Mensaje = ""
  StrArchivosAtach = ""
  Exit Sub
 End If
 ' ************************************************************************
 
 ' ************************************************************************
 ' Establece el Primer Contacto con el Servidor (Comando HELO)
 ' ************************************************************************
 Server.SMTPMailSocket.SendData ("HELO SMTPMail.By.PabloAmomo.Com" + vbCrLf)
 EsperarEvento "250", EmailTimeOut
 If EventoOK = 0 Then
  RespuestaEventoError = "02: No fue Posible Enviar el Comando de Inicialización al Servidor SMTP..."
  Server.SMTPMailSocket.Close
  Mensaje = ""
  StrArchivosAtach = ""
  Exit Sub
 End If
 ' ************************************************************************
 
 ' ************************************************************************
 ' Envia el Mail FROM
 ' ************************************************************************
 Server.SMTPMailSocket.SendData (MailDe)
 EsperarEvento "250", EmailTimeOut
 If EventoOK = 0 Then
  RespuestaEventoError = "03: No fue Posible Enviar el Comando [MAIL FROM:] al Servidor SMTP..."
  Server.SMTPMailSocket.Close
  Mensaje = ""
  StrArchivosAtach = ""
  Exit Sub
 End If
    
 ' ************************************************************************
 ' Envia el Mail RCPT TO:
 ' ************************************************************************
 Server.SMTPMailSocket.SendData (Casilla)
 EsperarEvento "250", EmailTimeOut
 If EventoOK = 0 Then
  RespuestaEventoError = "04: No fue Posible enviar el Comando [RCPT] al Servidor SMTP..."
  Server.SMTPMailSocket.Close
  Mensaje = ""
  StrArchivosAtach = ""
  Exit Sub
 End If
        
 ' ************************************************************************
 ' Envia los Datos (Mensaje) comando DATA
 ' ************************************************************************
  Server.SMTPMailSocket.SendData ("data" + vbCrLf)
  EsperarEvento "354", EmailTimeOut
  If EventoOK = 0 Then
   RespuestaEventoError = "05: No fue Posible enviar el comando [DATA] al Servidor SMTP..."
   Server.SMTPMailSocket.Close
   Mensaje = ""
   StrArchivosAtach = ""
   Exit Sub
  End If

    ' Envia el Encabezado del Mensaje...
    Server.SMTPMailSocket.SendData (EncabezadoMail)
    
    ' *****************************************************************************
    ' Envia el Mensaje...
    ' *****************************************************************************
     Dim VarLines, VarLine As Variant
         
     ' Parte el Mensaje en lineas delimitadas por el caracter chr$(13)
     VarLines = Split(Mensaje, vbCrLf)
     For Each VarLine In VarLines
       ' Envia el Mensaje con el Atach si es que tiene
       Server.SMTPMailSocket.SendData CStr(VarLine) & vbCrLf
     Next
     ' Liberar Memoria de Mensaje
     Mensaje = ""
     VarLine = ""
     VarLines = ""
    ' *****************************************************************************
    
    ' Concluye con el Envio
    Server.SMTPMailSocket.SendData vbCrLf
    Server.SMTPMailSocket.SendData ("." + vbCrLf)
    EsperarEvento "250", EmailTimeOut
    If EventoOK = 0 Then
     RespuestaEventoError = "06: No fue Posible Enviar el Mensaje al Servidor SMTP..."
     Server.SMTPMailSocket.Close
     Mensaje = ""
     Exit Sub
    End If

 ' ************************************************************************
 ' Se desconecta del Servidor SMTP
 ' ************************************************************************
 Server.SMTPMailSocket.SendData ("quit" + vbCrLf)
 EsperarEvento "221", EmailTimeOut
 If EventoOK = 0 Then
  RespuestaEventoError = "07: No fue Posible Realizar la Desconección del Servidor SMTP..."
  Server.SMTPMailSocket.Close
  Mensaje = ""
  Exit Sub
 End If

 ' ************************************************************************
 ' Cierra el Socket
 ' ************************************************************************
 Server.SMTPMailSocket.Close

Else
 RespuestaEventoError = "08: No fue posible Abrir el Socket para el Envio..."
End If

Exit Sub

SalirEnviando:
ErrorRecibido = Err.Description
If ErrorRecibido = "" Then ErrorRecibido = "No se Encontro Descripción para el Error..."
RespuestaEventoError = "10: Se produjo el Error: " & ErrorRecibido
Exit Sub

ErrorEnviando:
Server.SMTPMailSocket.Close
Mensaje = ""
StrArchivosAtach = ""
Resume SalirEnviando

End Sub
Private Sub EsperarEvento(Codigo As String, Tiempo As Integer)
Dim Reloj, Cronometro As Single

    ' El evento es  0 si MAL
    ' o             1 si OK
    
    Rem ****************************************
    Rem Permite verificar si el proceso
    Rem salio OK, o hubo algun problema
    Rem ****************************************
    EventoOK = 0
    Rem ****************************************
    
    Rem ****************************************
    Rem Se comienza a Contar para definir
    Rem TimeOuts
    Rem ****************************************
    Reloj = Timer
    Rem ****************************************
    
    Rem ****************************************
    Rem Se verifica hasta que la respuesta sea
    Rem distinta de 0, o que se supere el TimeOut
    Rem ****************************************
    While Len(RespuestaWinsockSMTP) = 0
        Cronometro = Timer - Reloj
        DoEvents
        If Cronometro > Tiempo Then
            Rem Time Out
            Exit Sub
        End If
    Wend
    Rem ****************************************
    
    Rem ****************************************
    Rem Cuando se recibe una respuesta se chequea
    Rem que sea la que se espera...
    Rem ****************************************
    While Left(RespuestaWinsockSMTP, 3) <> Codigo
        Cronometro = Timer - Reloj
        DoEvents
        If Cronometro > Tiempo Then
            Rem Mensaje de Error
            Exit Sub
        End If
    Wend
    Rem ****************************************

    Rem ****************************************
    Rem Permite verificar si el proceso
    Rem salio OK, o hubo algun problema
    Rem ****************************************
    EventoOK = 1
    Rem ****************************************

    RespuestaWinsockSMTP = ""
    
End Sub
Private Function CodificarArchivo(NombreDeArchivoCompleto As String) As String
On Error GoTo ErrorCodificacion
Dim TamanioArchivo, I, J, LineasACodificar As Long
Dim Resultado, NombreDeArchivo, BufferDatos As String
Dim LineaTemporal     As String

    ErrorCodArc = 0
    
    ' ******************************************************
    ' Levanta el Nombre del Archivo desde el Path completo
    ' ******************************************************
    NombreDeArchivo = Mid$(NombreDeArchivoCompleto, InStrRev(NombreDeArchivoCompleto, "\") + 1)
    
    ' ******************************************************
    ' Le agrega la primer marca de la codificacion
    ' ******************************************************
    ' ?: Resultado = "begin 664 " + NombreDeArchivo + vbLf
    Resultado = "begin 664 " + NombreDeArchivo + vbCrLf
    
    ' ******************************************************
    ' Levanta el Tamaño del Archivo, y verfica cuantas
    ' lineas de 45 caracteres debe codificar
    ' ******************************************************
    TamanioArchivo = FileLen(NombreDeArchivoCompleto)
    LineasACodificar = TamanioArchivo \ 45 + 1
    
    ' ******************************************************
    ' Arma un Buffer para utilizar en la Codificacion
    ' ******************************************************
    BufferDatos = Space(45)
        
    ' ************************************************
    ' Cierra el Archivo #1, por si las moscas
    ' ************************************************
    Close #1
    
    ' ************************************************
    ' Abre el Archivo y comienza la codificacion
    ' ************************************************
    Open NombreDeArchivoCompleto For Binary As #1
        For I = 1 To LineasACodificar
            If I = LineasACodificar Then
                BufferDatos = Space(TamanioArchivo Mod 45)
            End If
            Get #1, , BufferDatos
            LineaTemporal = Chr(Len(BufferDatos) + 32)
            '
            If I = LineasACodificar And (Len(BufferDatos) Mod 3) Then
                BufferDatos = BufferDatos + Space(3 - (Len(BufferDatos) Mod 3))
            End If
            
            For J = 1 To Len(BufferDatos) Step 3
                '1 byte
                LineaTemporal = LineaTemporal + Chr(Asc(Mid(BufferDatos, J, 1)) \ 4 + 32)
                '2 byte
                LineaTemporal = LineaTemporal + Chr((Asc(Mid(BufferDatos, J, 1)) Mod 4) * 16 _
                               + Asc(Mid(BufferDatos, J + 1, 1)) \ 16 + 32)
                '3 byte
                LineaTemporal = LineaTemporal + Chr((Asc(Mid(BufferDatos, J + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(BufferDatos, J + 2, 1)) \ 64 + 32)
                '4 byte
                LineaTemporal = LineaTemporal + Chr(Asc(Mid(BufferDatos, J + 2, 1)) Mod 64 + 32)
            Next J
            LineaTemporal = Replace(LineaTemporal, " ", "`")
            
            ' Agregar la nueva linea codificada
            ' ?: Resultado = Resultado + LineaTemporal + vbLf
            Resultado = Resultado + LineaTemporal + vbCrLf
            LineaTemporal = ""
        Next I
    Close #1

    ' Escribe las Marcas de Fin de Archivo
    ' ?: Resultado = Resultado & "`" & vbLf + "end" + vbLf
    Resultado = Resultado & "`" & vbCrLf + "end" + vbCrLf
    CodificarArchivo = Resultado
    Exit Function
    
SalirCodificarArchivo:
 Exit Function
 
ErrorCodificacion:
 CodificarArchivo = Err.Description
 ErrorCodArc = 1
 Resume SalirCodificarArchivo
 
End Function




