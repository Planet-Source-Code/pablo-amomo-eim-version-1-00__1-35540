Attribute VB_Name = "ComandosIntercambioDePaquete"
Option Explicit
Function ArmarPartes(UltimoIngreso As Date, Usuario As String, ParteNumero As Integer, ParteTotal As Integer, Datos As String, IDMensaje As String)
Dim MultiMensaje, EstaCompleto  As Boolean
Dim Mensaje, Comandos, MensajeCombinado, TipoDePaquete As String
Dim Handle, Contador, PaqueteEncontrado As Long


 ' **************************************************************
 ' Verifica si es un MultiChat
 ' **************************************************************
 If Mid$(Datos, 1, 1) = "M" Then
   ' **************************************************************
   ' Si no llego el Handle Sale
   ' **************************************************************
   If Not IsNumeric(Mid$(Datos, 2, 10)) Then Exit Function
   MultiMensaje = True
   Handle = Mid$(Datos, 2, 10)
   Mensaje = Mid$(Datos, 12)
 End If
 
 ' **************************************************************
 ' Verifica si es un UniMensaje
 ' **************************************************************
 If Mid$(Datos, 1, 1) = "U" Then
   MultiMensaje = False
   Handle = 0
   Mensaje = Mid$(Datos, 2)
 End If
 
 ' **************************************************************
 ' Busca el Mensaje Partido
 ' **************************************************************
 PaqueteEncontrado = 0
 If CantidadDePaquetesPartidos > 0 Then
  For Contador = 1 To CantidadDePaquetesPartidos
   ' Verifica si es Multichat
   If MultiMensaje = ArmadoDePaquete(Contador).MultiChat Then
     ' Verifica que el Handle Sea el Mismo
     If ArmadoDePaquete(Contador).Handle = Handle Then
       ' Verifica que el Usuario Sea el Mismo
       If Trim(UCase(ArmadoDePaquete(Contador).UsuarioEmisor)) = Trim(UCase(Usuario)) Then
         ' Verifica que la Cantidad de Paquetes sea la Misma
         If ArmadoDePaquete(Contador).PaquetesTotales = ParteTotal Then
           ' Verifica que el ID del Mensaje sea correcto
           If ArmadoDePaquete(Contador).IDMensaje = IDMensaje Then
            ''''
            ' Uf !!, si pasa todos estos chequeos quiere decir que esta todo OK
            ' y lo encontro...
            ''''
            PaqueteEncontrado = Contador
           End If
         End If
       End If
     End If
   End If
  Next
 End If
 
 ' **************************************************************
 ' Ahora si encontro el paquete lo modifica
 ' **************************************************************
 If PaqueteEncontrado <> 0 Then
   ArmadoDePaquete(PaqueteEncontrado).Datos(ParteNumero) = Mensaje
 End If
 
 ' **************************************************************
 ' Sino crea uno nuevo
 ' **************************************************************
 If PaqueteEncontrado = 0 Then
  ' Define la Cantidad de Paquetes
  CantidadDePaquetesPartidos = CantidadDePaquetesPartidos + 1
  ReDim Preserve ArmadoDePaquete(CantidadDePaquetesPartidos)
  ' Lleva a "" los datos de todos lo datos
  For Contador = 1 To 9
   ArmadoDePaquete(CantidadDePaquetesPartidos).Datos(Contador) = ""
  Next
  With ArmadoDePaquete(CantidadDePaquetesPartidos)
   .Datos(ParteNumero) = Mensaje
   .Handle = Handle
   .MultiChat = MultiMensaje
   .PaquetesTotales = ParteTotal
   .UltimoPaqueteRecibido = UltimoIngreso
   .UsuarioEmisor = Usuario
   .IDMensaje = IDMensaje
  End With
  ' Esto lo usa para no tener que hacer una funcion para verificar si esta
  ' completo...
  PaqueteEncontrado = CantidadDePaquetesPartidos
 End If
  
 ' **************************************************************
 ' Verifica si el paquete esta completo
 ' **************************************************************
 EstaCompleto = True
 For Contador = 1 To ParteTotal
  If ArmadoDePaquete(CantidadDePaquetesPartidos).Datos(Contador) = "" Then
    ' Como encontro uno vacio,por lo cual no esta completo sale...
    EstaCompleto = False
    Exit For
   Else
    ' Si tiene algo comianza a combinar el Mensaje
    MensajeCombinado = MensajeCombinado & ArmadoDePaquete(CantidadDePaquetesPartidos).Datos(Contador)
  End If
 Next
 
 ' **************************************************************
 ' Si llega aca es por que esta completo
 ' **************************************************************
 ' Define el Comando que antecede al Mensaje
 If EstaCompleto Then
  If ArmadoDePaquete(CantidadDePaquetesPartidos).MultiChat Then
    ' MultiChat
    MensajeCombinado = "M" & CompletarCadena(CStr(ArmadoDePaquete(CantidadDePaquetesPartidos).Handle), 10, "I", "0") & MensajeCombinado
   Else
    ' Uni Usuario
    MensajeCombinado = "U" & MensajeCombinado
  End If
  ' Envia el Mensaje Completo para ser procesado
  ProcesarMensaje Usuario, CStr(MensajeCombinado)
  ' Una vez procesado lo borra
  ' Si es el ultimo simplemente baja 1 la  cantidad
  If CantidadDePaquetesPartidos = PaqueteEncontrado Then
    CantidadDePaquetesPartidos = CantidadDePaquetesPartidos - 1
   Else
    ' Sino pasa el Ultimo al actual y baja la cantidad en 1
    ArmadoDePaquete(CantidadDePaquetesPartidos) = ArmadoDePaquete(PaqueteEncontrado)
    CantidadDePaquetesPartidos = CantidadDePaquetesPartidos - 1
  End If
 End If
 
 ' **************************************************************
 ' Si el paquete tiene mas de 60 segundos lo borra, y pone el ultimo
 ' en primer lugar... Salvo que sea el Ultimo
 ' **************************************************************
 If CantidadDePaquetesPartidos > 0 Then
  Contador = 0
  Do
   Contador = Contador + 1
   ' Verifica que existan paquetes que analizar
   If Contador > CantidadDePaquetesPartidos Then Exit Do
   ' Pasaron 60 Segundos?
   If DateDiff("s", ArmadoDePaquete(Contador).UltimoPaqueteRecibido, UltimoIngreso) >= 60 Then
    ' Si es el Ultimo lo borra
    If Contador = CantidadDePaquetesPartidos Then
      CantidadDePaquetesPartidos = CantidadDePaquetesPartidos - 1
      ' Si es el ultimo paque lo Mueve y Sale
      Exit Do
     Else
      ArmadoDePaquete(Contador) = ArmadoDePaquete(CantidadDePaquetesPartidos)
      ' Hace esto para que cuando mueva, lo vuelva a escanear
      Contador = Contador - 1
      ' Como el ultimo paquete se paso al actual, se elimina el ultimo
      CantidadDePaquetesPartidos = CantidadDePaquetesPartidos - 1
      ' Si era el Ultimo Sale
      If CantidadDePaquetesPartidos = 0 Then Exit Do
    End If
   End If
  Loop
 End If
 
End Function
Function IntercambioDePaquete(Datos As Variant)
Dim Comando, ComandoGeneral, CadenaTemp, VentanaID As String
Dim Resto, UsuarioEmisor, Ventana, Resp As String
Dim Respuesta, Contador, Cantidad, Bandera As Integer
Dim Usuario, HoraYFecha, Mensaje As String

 ' **************************************************************
 ' Formato del Paquete:
 '          1       (1-2)   Respuesta del Server o Paquete del Usuario
 '              1:
 '                      1-  Respuesta Del Server
 '                                          0 Ok
 '                                          1 No Conectado
 '                                          2 No Disponible
 '                                          3 Usuario No Existe
 '                      16- Usuario al Que se le evniava el Mensaje
 '              2:
 '                      16- Emisor del Mensaje
 '                      Resto Mensaje
  
 ' Comandos enViados en RestodelMensaje
 '          2. Solicitud de Unirse a Multicaht
 '                  Composicion: 1 Comando, 1 Cantidad de Amigos, Amigos (Formato 16 caracteres)
 '          3. Confirmacion de Union a Multichat
 '                  Composicion: 0 No, 1 Ok
 ' **************************************************************
 
 ComandoGeneral = Mid$(Datos, 1, 1)
 UsuarioEmisor = Trim(Mid$(Datos, 2, 16))
 Comando = Trim(Mid$(Datos, 18, 1))
 Resto = Mid$(Datos, 19)
 
 ' **************************************************************
 ' Comando alto Nivel
 ' **************************************************************
 ' Aca decide si es un Paquete de respuesta del Servidor o no
 If ComandoGeneral = "1" Then Exit Function
 
 ' **************************************************************
 ' 2. Recibe la Solicitud de Unirse a un Multichat
 ' **************************************************************
 If Comando = "2" Then
  ' **************************************************************
  ' Verificar que el Amigo Emisor Exista en el Listado, Sino,
  ' pregunta si lo quiere agregar como amigo...
  ' **************************************************************
  Resp = Varios.VerificarUsuarioEnListadoAmigos(CStr(UsuarioEmisor))
  If UCase(Trim(Resp)) = "NO" Then
   EnviarPaqueteTCP ("3" & CompletarCadena(CStr(UsuarioEmisor), 16, "D", " ") & "30")
   MostrarMSGBox MensajeRecurso(464), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Exit Function
  End If
  ' El Amigo [ % ] le solicita Unirse a su MultiChat... ¿Desea Aceptar?
  Respuesta = MostrarMSGBox(MensajeRecurso(134) & UsuarioEmisor & MensajeRecurso(306), vbYesNo, "vbQuestion", Configuracion.TituloVentanas, False)
  '''' ACA !!!!!
  If Respuesta = vbYes Then
    ' **************************************************************
    ' Crea el MultiChat
    ' **************************************************************
    ' Define la Cantidad de Amigos en el Multichat
    Cantidad = CInt(Mid$(Resto, 1, 1))
    If Cantidad = 0 Then Cantidad = 10
    ' Crea la Ventana contra el usuario
    Respuesta = CrearVentanaMensaje(CStr(UsuarioEmisor))
        
    ' **************************************************************
    ' Envia la Confirmacion de la Union a MultiChat
    ' **************************************************************
    VentanaID = CompletarCadena(CStr(Forms(Respuesta).hwnd), 10, "I", "0")
    ' Aca envia el Hwnd como Ventna ID para que cuando viene un Listado, puedo
    ' ubicar el Form correcto...
    CadenaTemp = "3" & CompletarCadena(Trim(UsuarioEmisor), 16, "D", " ") & "31" & VentanaID
    EnviarPaqueteTCP CStr(CadenaTemp)
    
    ' **************************************************************
    ' Carga los MultiChat
    ' **************************************************************
    Resto = Mid$(Resto, 2)
    For Contador = 1 To Cantidad
     ' Carga el Nombre del Responsable del Multichat...
     Forms(Respuesta).ResponsableMultichatUsuario = UsuarioEmisor
     If Contador = 1 Then
      Forms(Respuesta).ResponsableMultichatVentanaID = Mid$(Resto, 1 + (26 * (Contador - 1)) + 16, 10)
     End If
     ''
     ' Solo agrega los Que son distintos a usuario emisor, debido a que con el
     ' usuario emisor se crea la ventna, y automaticamente pasa a ser un usuauio
     ' del MultiChat..., sino solo le pone el ID de Ventana al ya existente amigo...
     Respuesta = BuscarVentanaHandle(CLng(VentanaID))
     If UCase(Trim(Mid$(Resto, 1 + (26 * (Contador - 1)), 16))) <> UCase(Trim(UsuarioEmisor)) Then
       Forms(Respuesta).TratarAmigosEnChat "Agregar", Mid$(Resto, 1 + (26 * (Contador - 1)), 16), Mid$(Resto, 1 + (26 * (Contador - 1)) + 16, 10)
      Else
       ' Aca le Modifica el Ventana ID del Amigo que Envia la Solicitud
       Forms(Respuesta).TratarAmigosEnChat "ModificarID", CStr(Mid$(Resto, 1 + (26 * (Contador - 1)), 16)), CStr(Mid$(Resto, 1 + (26 * (Contador - 1)) + 16, 10))
     End If
    Next
   Else
    ' Envia la Negacion de la Incorporacion de Multichat
    EnviarPaqueteTCP ("3" & CompletarCadena(CStr(UsuarioEmisor), 16, "D", " ") & "30")
  End If
  Exit Function
 End If
  
 ' **************************************************************
 ' 3. Confirma o no La solicitud de Multichat
 ' **************************************************************
 If Comando = "3" Then
  For Contador = 1 To Forms.Count - 1
   ' Busca los Formularios de Mensaje
   If Forms(Contador).FormularioNombre = "Mensajes" Then
    ' Esta es la Ventana
    Respuesta = Forms(Contador).BuscarUsuarioPendiente(Trim(UsuarioEmisor))
    If Respuesta = -1 Then
     Forms(Contador).AgregarChayMultiusuarioPendiente "Confirmar", CStr(Trim(UsuarioEmisor)), CInt(Mid$(Resto, 1, 1)), Trim(Mid$(Resto, 2, 10))
    End If
   End If
  Next
 End If
  
 ' **************************************************************
 ' 4. Nuevo Listado de Amigos MultiChat
 ' **************************************************************
 If Comando = "4" Then
  ' **************************************************************
  ' Verifica el Largo del Paquete
  ' **************************************************************
  If Len(Resto) < 11 Then Exit Function
  ' **************************************************************
  ' Separa el Paquete
  ' **************************************************************
  Ventana = Mid$(Resto, 1, 10)
  Resto = Mid$(Resto, 11)
  ' **************************************************************
  ' Verifica que el Numero de Vantana sea correcto (Es decir que sea un Numero)
  ' **************************************************************
  If Not IsNumeric(Ventana) Then Exit Function
  ' **************************************************************
  ' Ubica la Ventana con el Handle recibido
  ' **************************************************************
  Respuesta = 0
  Respuesta = BuscarVentanaHandle(CLng(Ventana))
  'For Contador = 1 To Forms.Count - 1
  ' If CompletarCadena(CStr(Forms(Contador).hwnd), 10, "I", "0") = Ventana Then
  '  Respuesta = Contador
  '  Exit For
  ' End If
  'Next
  
  ' **************************************************************
  ' Si no encontro la Ventana Sale...
  ' **************************************************************
  If Respuesta = 0 Or Respuesta = "" Then Exit Function
  
  ' **************************************************************
  ' Agrega los Amigos Al MultiChat
  ' **************************************************************
  ' Lleva a Cero Los Amigos...
  Forms(Respuesta).CantidadDeAmigosEnChat = 0
  Forms(Respuesta).CantidadDeAmigosEnChatPendiente = 0
  Cantidad = CInt(Mid$(Resto, 1, 1))
  If Cantidad = 0 Then Cantidad = 10
  
  ' **************************************************************
  ' Carga los MultiChat
  ' **************************************************************
  Resto = Mid$(Resto, 2)
  Bandera = 0
  For Contador = 1 To Cantidad
   If UCase(Trim(Mid$(Resto, 1 + (26 * (Contador - 1)), 16))) <> UCase(Trim(Configuracion.IDAliasUsuario)) Or Trim(Mid$(Resto, 1 + (26 * (Contador - 1)) + 16, 10)) <> CompletarCadena(Forms(Respuesta).hwnd, 10, "I", "0") Then
     Forms(Respuesta).TratarAmigosEnChat "Agregar", Trim(Mid$(Resto, 1 + (26 * (Contador - 1)), 16)), Trim(Mid$(Resto, 1 + (26 * (Contador - 1)) + 16, 10))
    Else
     ' Define que sigue estando en el MulTichat
     Bandera = 1
   End If
  Next
    
  ' **************************************************************
  ' Verifica que no haya sifo eliminado del MultiChat
  ' **************************************************************
  ' El usuario fue eliminado del MultiChat
  If Bandera = 0 Then
   ' El Amigo [ % ] lo Ha Eliminado de su * MultiChat *...
   MostrarMSGBox MensajeRecurso(134) & UsuarioEmisor & MensajeRecurso(307), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   With Forms(Respuesta)
    .CantidadDeAmigosEnChat = 0
    .CantidadDeAmigosEnChatPendiente = 0
    .TratarAmigosEnChat "Agregar", Trim(UsuarioEmisor)
   End With
   Exit Function
  End If
 End If
 
 ' **************************************************************
 ' 5. Mensaje offline
 ' **************************************************************
 If Comando = "5" Then
  ' **************************************************************
  ' Formato del Paquete:
  '  1  Carataer Comando Adicional (1 Confirmacion / 2 Mensaje)
  '  16 Caracteres el Nombre del User Emisor
  '  19 Caracteres FechayHora (Formato) HH:MM:SS_DD/MM/AAAA
  '  Resto Mensaje...
  ' **************************************************************
  If Len(Resto) < 2 Then Exit Function ' Paquete invalidp
  ' Saca el Comando adicional
  Dim ComandoAdicional As String
  ComandoAdicional = Trim(Mid$(Resto, 1, 1))
  Resto = Trim(Mid$(Resto, 2))
  ' **************************************************************
  ' Define si es una Confirmacion o un Mensaje...
  ' **************************************************************
  If ComandoAdicional = 1 Then
   If Len(Resto) < 17 Then Exit Function ' Paquete Invalido
   Usuario = Trim(Mid$(Resto, 1, 16))
   Mensaje = Mid$(Resto, 17, 1) ' En Realidad esto es la Respuesta...
   Contador = Varios.BuscarFormularioNombre("MensajeOffline", CStr(Usuario))
   If Contador <> -1 Then
    ' Lo Encontro...
    Forms(Contador).RecibidoOk = CInt(Mensaje)
   End If
    ' Si no loencuentra descarta la devolucion...
   Exit Function
  End If
  ' **************************************************************
  ' Define que es un Mensaje
  ' **************************************************************
  If ComandoAdicional = 2 Then
   If Len(Resto) < 36 Then Exit Function ' Paquete Invalido
   ' **************************************************************
   ' Descompone el Paquete
   ' **************************************************************
   Dim EstadoUsuario As String
   Usuario = Trim(Mid$(Resto, 1, 16))
   HoraYFecha = Trim(Mid$(Resto, 17, 19))
   Mensaje = Mid$(Resto, 36)
   ' Si el Mensaje empieza con chr$(0) & Chr$(0) => es un Amigo que incorporo...
   If Len(Mensaje) >= 16 Then
    Comando = Chr$(0) & Chr$(0)
    If Mid$(Mensaje, 1, 2) = Comando Then
     ProcesarAvisoDeIncorporacion Mensaje, CStr(Usuario)
     Exit Function
    End If
   End If
   VerificarUsuarioEnListadoAmigos CStr(Usuario)
   Varios.AgregarMensajesPendientes CStr(Usuario), CStr(Mensaje), CStr(HoraYFecha)
   Exit Function
  End If
 End If
 
 ' **************************************************************
 ' 6. Alguien me incorporo comom Amigo...
 ' **************************************************************
 'If Comando = "6" Then
 ' If Len(Resto) < 1 Then Exit Function ' Paquete Invalido...
 ' Dim Texto As String
 ' Mensaje = Trim(Resto)
 ' MsgBox "Hola, soy " & UsuarioEmisor & " [" & Mensaje & "] y acabo de Incorporarlo ami Listado de Amigos..."
 ' Exit Function
 'End If
 
 ' **************************************************************
 ' 7. Avisa algun Evento de un usuario Determinado...
 ' **************************************************************
 If Comando = "7" Then
  If Len(Resto) < 10 Then Exit Function ' Paquete Invalido...
  Dim Handle, Evento, MasDatos As String
  Dim Posicion As Integer
  Usuario = UsuarioEmisor
  Handle = Trim(Mid$(Resto, 1, 10))
  Evento = Mid$(Resto, 11, 1)
  MasDatos = ""
  If Len(Resto) > 11 Then
   MasDatos = Mid$(Resto, 12)
  End If
  Respuesta = 0
  For Contador = 1 To Forms.Count - 1
   ' Revisa las Ventanas de Formulario
   If Forms(Contador).FormularioNombre = "Mensajes" Then
    If Handle = "0000000000" Then
     ' Lo Busca por Usuario
      Posicion = Forms(Contador).TratarAmigosEnChat("Buscar", CStr(Usuario))
      If Posicion = 1 Then
       ' Es el Primer Amigo, Entonces no es un Multichat...
       Respuesta = Contador
       Exit For
      End If
     Else
     ' Lo Busca por Handle (Es decir Multichat)
      If CompletarCadena(CStr(Forms(Contador).hwnd), 10, "I", "0") = Handle Then
       Respuesta = Contador
       Exit For
      End If
    End If
   End If
  Next
  If Respuesta = 0 And Evento <> "4" Then Exit Function ' No Encontro la Ventana...
  Select Case UCase(Evento)
   Case "1" ' Escribiendo un Mensaje...
    Forms(Respuesta).AgregarEventoDelUsuario MensajeRecurso(134) & Trim(Usuario) & MensajeRecurso(474)
   Case "2" ' Salio del Multichat...
    Forms(Respuesta).AgregarEventoDelUsuario MensajeRecurso(134) & Trim(Usuario) & MensajeRecurso(475)
    Forms(Respuesta).FunctionAlguienDejoElMultichat Trim(Usuario)
   Case "3" ' El Responsable Cancelo el MultiChat...
    Forms(Respuesta).AgregarEventoDelUsuario MensajeRecurso(134) & Trim(Usuario) & MensajeRecurso(476)
    Forms(Respuesta).CancelarMultichat MensajeRecurso(134) & Trim(Usuario) & MensajeRecurso(476)
   Case "4" ' Avisa que un Usuario X, ya no es parte del Multichat...
    Dim EventoNombre, EventoHandle, EventoMensaje, PaqueteEnviar As String
    If Len(MasDatos) < 27 Then Exit Function
    EventoHandle = Mid$(MasDatos, 1, 10)
    EventoNombre = Mid$(MasDatos, 11, 16)
    EventoMensaje = Mid$(MasDatos, 27)
    If Not IsNumeric(EventoHandle) Then Exit Function
    For Contador = 0 To Forms.Count - 1
     If UCase(Forms(Contador).FormularioNombre) = UCase("mensajes") Then
      Respuesta = Forms(Contador).TratarAmigosEnChat("BuscarHandleYNombre", CStr(Usuario), CStr(EventoHandle))
      If Respuesta <> 0 Then
       Forms(Contador).AgregarMensaje CStr(Usuario), CStr(EventoMensaje), CStr(Time & "_" & Date)
       ' Aca debe hacer el Sacado del Usuario...
       ' Si es el Owner lo sca, sino manda un Mensaje para que el Owner lo
       ' Saque
       If Forms(Contador).ResponsableMultiChat Then
         ' Es el Dueño...
         Forms(Contador).TratarAmigosEnChat "Sacar", CStr(Usuario)
         Forms(Contador).CargarAmigosMultiChat
         Forms(Contador).EnviarListadoDeMultiChat
        Else
         ' Si no es el Du#o envia el Evento 5 con el usuario a Borrar...
         Varios.EnviarBorradousuario Trim(Forms(Contador).ResponsableMultichatUsuario), Trim(Forms(Contador).ResponsableMultichatVentanaID), Trim(Usuario), CStr(Trim(EventoHandle))
         
         'PaqueteEnviar = "43" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
                        "M" & "01" & _
                        CompletarCadena(Trim(Forms(Contador).ResponsableMultichatUsuario), 16, "D", " ") & _
                        CompletarCadena(Trim(Forms(Contador).ResponsableMultichatVentanaID), 10, "I", "0") & "5" & _
                        CompletarCadena(Trim(EventoHandle), 10, "I", "0") & CompletarCadena(Trim(Usuario), 16, "D", " ")

         
         
         'PaqueteEnviar = "3" & CompletarCadena(Trim(Forms(Contador).ResponsableMultichatUsuario), 16, "D", " ") & _
         '               "5" & CompletarCadena(Forms(Contador).ResponsableMultichatVentanaID, 10, "I", "0") & _
         '               "01" & CompletarCadena(Trim(EventoHandle), 10, "I", "0") & CompletarCadena(Trim(Usuario), 16, "D", " ")
         'EnviarPaqueteTCP PaqueteEnviar
       End If
      End If
     End If
    Next
      
   Case "5" ' Se Solicita borrar un amigo del Multichat...
    If Len(MasDatos) < 11 Then Exit Function
    EventoHandle = Mid$(MasDatos, 1, 10)
    EventoNombre = Mid$(MasDatos, 11)
    If Not IsNumeric(EventoHandle) Then Exit Function
    Resp = Forms(Contador).TratarAmigosEnChat("BuscarHandleYNombre", CStr(EventoNombre), CStr(EventoHandle))
    If Resp <> 0 Then
     Dim VentanaIDResponsable As Long ' ESTA LINEA HAY QUE BORRARLA...
     Dim VentanaIDResp As String
     VentanaIDResp = Forms(Respuesta).ResponsableMultichatVentanaID
     If VentanaIDResp = "" Or Not IsNumeric(VentanaIDResp) Then VentanaIDResp = "0"
     If UCase(Trim(Forms(Respuesta).ResponsableMultichatUsuario)) = UCase(Trim(EventoNombre)) And CLng(VentanaIDResp) = CLng(EventoHandle) Then
      Exit Function
     End If
     Forms(Respuesta).TratarAmigosEnChat "Sacar", CStr(EventoNombre), CStr(EventoHandle)
     Forms(Respuesta).CargarAmigosMultiChat
     Forms(Respuesta).EnviarListadoDeMultiChat
    End If
   End Select
  Exit Function
 End If
 
End Function
Function ProcesarAvisoDeIncorporacion(Mensaje As String, Usuario As String)
Dim NuevoAmigoAlias, NuevoAmigoNombre As String
Dim UsuarioAlias, EstadoAmigoTexto, Sexo As String
Dim Estadoamigo, Contador, Respuesta As Integer
Dim Usuarioexiste As Boolean
Dim NuevoAmigoBandera As Boolean
     
 ' **************************************************************
 ' Carga los Datos del Amigo...
 ' **************************************************************
 NuevoAmigoAlias = Mid$(Mensaje, 3, 16)
 NuevoAmigoNombre = Mid$(Mensaje, 19, 50)
 
 ' **************************************************************
 ' Primero verifica si el Amigo que envio existe en la lista...
 ' **************************************************************
 NuevoAmigoBandera = False
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(NuevoAmigoAlias)) Then
   NuevoAmigoBandera = True
   Exit For
  End If
 Next
 Mensaje = MensajeRecurso(134) & Trim(NuevoAmigoAlias) & MensajeRecurso(465)
 
 ' **************************************************************
 ' Muestra el Mensage Box correspondiente...
 ' **************************************************************
 If NuevoAmigoBandera = False Then
   ''''''' Si ya le pregunto y cancelo que no le pregunte... No Pregunta mas...
   'If Varios.VolverAPreguntar(Trim(NuevoAmigoAlias)) = True Then Exit Function
   
   Mensaje = Mensaje & Chr$(13) & MensajeRecurso(466)
   Respuesta = MostrarMSGBox(Mensaje, vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbYes Then
    ' **************************************************************
    ' Lo agrega como amigo...
    ' **************************************************************
    Respuesta = SolicitarAgregarAmigo(Usuario)
    UsuarioAlias = Trim(Variables.RespuestaAmigoACrearAlias)
    Estadoamigo = Variables.RespuestaAmigoACrearEstado
    EstadoAmigoTexto = Trim(Variables.RespuestaAmigoACrearEstadoTexto)
    Sexo = Trim(Variables.RespuestaAmigoACrearSexo)
    Usuarioexiste = True
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
   End If
  Else
   MostrarMSGBox Mensaje, vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 End If


End Function
Function IntercambioDeMensaje(Datos As Variant) As String
Dim Comando, Resto, Usuario, Mensaje, CadenaTemp, IDMensaje, Adicional As String
Dim Contador, Respuesta, VentanaID As Integer
Dim HoraYFecha As String
Dim ParteActual, ParteTotal As String

 ' **************************************************************
 ' Paquete enviado al Emisor:
 '      0 Enviado OK
 '      1 No Conectado
 '      2 No Disponible
 '      3 Usuario No Existe
 '      M Mensaje Enviado
 '              16 Caracteres Emiso, Resto Mensaje
 ' **************************************************************
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Datos) < 17 Then
  ' Paquete Invalido...
  Exit Function
 End If
 
 ' **************************************************************
 ' Descompone el Paquete
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 Resto = Trim(Mid$(Datos, 2, 16))
 If Len(Datos) > 17 Then
   Adicional = Trim(Mid$(Datos, 18))
  Else
   Adicional = ""
 End If
 
 ' **************************************************************
 ' Procesa la Info de como llego el Paquete
 ' **************************************************************
 Select Case Comando
  Case "0", "1", "2", "3"
   If Adicional <> "" Then
     PonerRespuestaRecepcion Trim(Resto), CInt(Comando), Adicional
    Else
     PonerRespuestaRecepcion Trim(Resto), CInt(Comando)
   End If
  Exit Function
 End Select
 
 ' **************************************************************
 ' Verifica que sea un Mensaje
 ' **************************************************************
 If UCase(Comando) <> "M" Then Exit Function
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Datos) < 23 Then Exit Function
 
 ' **************************************************************
 ' Separa el Paquete
 ' **************************************************************
 Usuario = Trim(Mid$(Datos, 2, 16))
 ParteActual = Mid$(Datos, 19, 1)
 ParteTotal = Mid$(Datos, 20, 1)
 IDMensaje = Mid$(Datos, 21, 2)
 Mensaje = Mid$(Datos, 18, 1) & Mid$(Datos, 23)
 
 ' **************************************************************
 ' Verifica que Parte que se recibe y el total sean numericos
 ' **************************************************************
 If Not IsNumeric(ParteActual) Or Not IsNumeric(ParteTotal) Then
  Exit Function
 End If
 
 ' **************************************************************
 ' Si es un Mensaje de una Parte sola lo manda directo sino lo
 ' procesa...
 ' **************************************************************
 If CInt(ParteActual) = 1 And CInt(ParteTotal) = 1 Then
   ' Solo una Parte y es la parte 1
   ProcesarMensaje CStr(Usuario), CStr(Mensaje)
  Else
   ' Junta las partes, si esta completo...
   ArmarPartes Time, CStr(Usuario), CInt(ParteActual), CInt(ParteTotal), CStr(Mensaje), CStr(IDMensaje)
 End If
 
End Function
Public Function PonerRespuestaRecepcion(Usuario As String, Recepcion As Integer, Optional RespuestaAdicional As String)
Dim Contador As Integer
Dim Respuesta As Integer

   For Contador = 0 To Forms.Count - 1
    ' Ubica los Formularios de Mensaje
    If Forms(Contador).FormularioNombre = "Mensajes" Then
     Respuesta = Forms(Contador).TratarAmigosEnChat("Buscar", Usuario)
     ' Esta es la Ventana Correcta
     If Respuesta = 1 Then
      Forms(Contador).RecibidoOk = CInt(Recepcion)
      If Not IsNull(RespuestaAdicional) Then
       Forms(Contador).RecibidoOkRespuesta = RespuestaAdicional
      End If
     End If
    End If
   Next
     
End Function
Public Function BuscarVentana(Usuario As String, Optional MultiChat As String)
Dim Contador As Integer
Dim Respuesta As Integer

   For Contador = 0 To Forms.Count - 1
    ' Ubica los Formularios de Mensaje
    If Forms(Contador).FormularioNombre = "Mensajes" Then
     Respuesta = Forms(Contador).TratarAmigosEnChat("Buscar", Usuario)
     ' Esta es la Ventana Correcta
     If Respuesta = 1 Then
      ' Verifica si se pide busqueda o no de Multichat
      If Trim(MultiChat) <> "" Then
        ' Verifica la Peticion...
        If MultiChat = "Si" And Forms(Contador).CantidadDeAmigosEnChat > 1 Then
         BuscarVentana = Contador
         Exit Function
        End If
        If MultiChat = "No" And Forms(Contador).CantidadDeAmigosEnChat = 1 Then
         BuscarVentana = Contador
         Exit Function
        End If
       Else
        BuscarVentana = Contador
        Exit Function
      End If
     End If
    End If
   Next
   
   ' **************************************************************
   ' No la Encontro...
   ' **************************************************************
   BuscarVentana = 0
   
End Function
Public Function BuscarVentanaHandle(Handle As Long) As Integer
Dim Contador As Integer
Dim Respuesta As Integer

   ' **************************************************************
   ' Busca el Formulario Correspondiente
   ' **************************************************************
   For Contador = 0 To Forms.Count - 1
    ' Ubica El Formulario Correcto
    If Forms(Contador).hwnd = Handle Then
     BuscarVentanaHandle = Contador
     Exit Function
    End If
   Next
   
   ' **************************************************************
   ' No la Encontro...
   ' **************************************************************
   BuscarVentanaHandle = 0
   
End Function
Sub ProcesarMensaje(Usuario As String, Mensaje As String, Optional ProcesarIgual As String)
Dim Respuesta, Estadoamigo, Contador As Integer
Dim CadenaTemp, Sexo, Comando, MensajeFinal, EstadoAmigoTexto As String
Dim IDVentana As Long
Dim Bandera, Usuarioexiste As Boolean
 
 ' **************************************************************
 ' Verificar si el usuario emisor existe en el listado de Amigos
 ' **************************************************************
 Varios.VerificarUsuarioEnListadoAmigos Usuario
 

 ' **************************************************************
 ' Verifica que no sea un Usuario Bloqueado, Si lo
 ' es sale sin hacer nada...
 ' **************************************************************
 For Contador = 1 To Variables.UsuarioBloqueadoCantidad
  If UCase(Trim(Usuario)) = UCase(Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)) Then
   Exit Sub
  End If
 Next

 ' **************************************************************
 ' Valida el Largo del Mensaje
 ' **************************************************************
 If Len(Mensaje) < 1 Then
  ' Largo Invalido...
  Exit Sub
 End If
 
 ' **************************************************************
 ' Separa el Paquete
 ' **************************************************************
 ' Comando: U=UniUsuario        M=MultiChat
 Comando = Mid$(Mensaje, 1, 1)
 ' Saque el IDVentana si es un Mensaje Multichat
 If UCase(Comando) = "M" Then ' MultiChat
   ' 1. Valida el Largo del Paquete
   If Len(Mensaje) < 12 Then Exit Sub ' Tiene en Cuenta que el Mensaje tenga 1 por lo menos...
   ' 2. Verifica que el ID sea un Numero Long
   If Not IsNumeric(Mid$(Mensaje, 2, 10)) Then Exit Sub
   IDVentana = CLng(Mid$(Mensaje, 2, 10))
   MensajeFinal = Mid$(Mensaje, 12)
  Else
   ' Si Procesar Igual <> "" entonces no le da bola al largo
   If UCase(Trim(ProcesarIgual)) <> "SI" Then
     ' 1. Valida el Largo del Paquete
     If Len(Mensaje) < 2 Then Exit Sub ' Tiene en Cuenta que el Mensaje tenga 1 por lo menos...
     MensajeFinal = Mid$(Mensaje, 2)
    Else
     MensajeFinal = ""
   End If
 End If
 ' **************************************************************
  
 ' **************************************************************
 ' Busca si Existe Ventana contra dicho Usuario, Sino, lo pone
 ' como mensaje pendiente... (Solo para mensajes "U")
 ' **************************************************************
 If UCase(Comando) = "U" Then
  Respuesta = BuscarVentana(CStr(Usuario), "No")
  If Respuesta = 0 Then
   ' Pone el Mensaje en Pendiente o Muestra la
   ' Ventana segun si Cliente esta o no Visible...
   If Cliente.Visible = False Then
     ' Agrega el Mensaje como Pendiente
     AgregarMensajesPendientes CStr(Usuario), CStr(MensajeFinal)
     Exit Sub
    Else
     ' Si cliente no esta Minimizado, es decir
     ' No visible abre la Ventana Automaticamente, siempre y cuando
     ' no exista una inputbox o msgbox abierto en modo modal
     If Varios.BuscarFormularioNombre("MensajesBox") <> -1 Or Varios.BuscarFormularioNombre("IngresoBox") <> -1 Then
       AgregarMensajesPendientes CStr(Usuario), CStr(MensajeFinal)
       Exit Sub
      Else
       Dim Handle As Long
       Handle = CrearVentanaMensaje(CStr(Usuario))
       Respuesta = BuscarVentanaHandle(Handle) ' BuscarVentana(CStr(Usuario))
     End If
   End If
  End If
 End If
 
 ' **************************************************************
 ' Pone el Mensaje en una O todas las Ventanas - Mensaje "U"
 ' **************************************************************
 Dim PusoElMensaje As Boolean
 PusoElMensaje = False
 For Contador = 0 To Forms.Count - 1
  ' Busca las Ventanas de Mensaje
  If Forms(Contador).FormularioNombre = "Mensajes" Then
   CadenaTemp = Time & "_" & Date
   Select Case Comando
    Case "U"
     ' Chequea la Ventana verificando que exista el usuario
     Respuesta = Forms(Contador).TratarAmigosEnChat("Buscar", CStr(Usuario))
     ' Verifica que el Usuario exista, y no sea un MultiChat...
     If Respuesta = 1 And Forms(Contador).CantidadDeAmigosEnChat = 1 Then
      Forms(Contador).AgregarMensaje CStr(Usuario), CStr(MensajeFinal), CStr(CadenaTemp)
      PusoElMensaje = True
     End If
    Case "M"
     ' Verifica la Ventana Comparandolo con el Ventana ID recibido...
     If Forms(Contador).hwnd = IDVentana Then
      Forms(Contador).AgregarMensaje CStr(Usuario), CStr(MensajeFinal), CStr(CadenaTemp)
      PusoElMensaje = True
     End If
   End Select
  End If
 Next
 
 ' **************************************************************
 ' Avisa que actualmente no se encuentra incluido en el Multichat que llego...
 ' **************************************************************
 If PusoElMensaje = False Then
  If Comando = "M" Then
   Respuesta = Varios.MostrarMSGBox(MensajeRecurso(134) & Usuario & MensajeRecurso(462), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbYes Then
    MensajeDeNoEnMultichat Usuario, CStr(IDVentana)
   End If
  End If
 End If
  
End Sub
Public Sub MensajeDeNoEnMultichat(Amigo As String, IDVentana As String)
Dim RichTemporalX As Object
Dim Contador As Integer
Dim ComandoAdicional As String
      
    Set RichTemporalX = CreateObject("RICHTEXT.RichtextCtrl.1")
    ' *********************************************************************
    ' Define el Mensaje Correspondiente...
    ' *********************************************************************
    RichTemporalX.SelStart = Len(RichTemporalX.Text) + 1
    ' Pone la Cara...
    Clipboard.Clear
    Clipboard.SetData Cliente.ImagenCaras.ListImages(3).Picture
    SendMessage RichTemporalX.hwnd, &H302, 0, 0
    Clipboard.Clear
    ' Pone el Texto Correspondiente...
    RichTemporalX.SelRTF = "{{{\colortbl ;\red128\green128\blue128;}" & _
                             "{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & _
                             "\viewkind4\uc1\pard\cf1\b\fs16 " & _
                             "\b    " & MensajeRecurso(461) & "\cf0\b0\f1\fs17}}"
    
 
 ' **************************************************************
 ' Enviar Evento al Usuario Emisor...
 ' **************************************************************
 'ComandoAdicional = ComandoAdicional & CompletarCadena(Amigo, 16, "D", " ") & CompletarCadena(IDVentana, 10, "I", "0") & "4" & RichTemporalX.TextRTF
 'ComandoAdicional = "M" & CompletarCadena(CStr(1), 2, "I", "0") & ComandoAdicional
 'EnviarPaqueteTCP "47" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
 '                 ComandoAdicional

 ComandoAdicional = "M" & CompletarCadena("1", 2, "I", "0")
 ComandoAdicional = ComandoAdicional & CompletarCadena(Amigo, 16, "D", " ") & CompletarCadena(IDVentana, 10, "I", "0") & "4"
 ComandoAdicional = ComandoAdicional & CompletarCadena(CStr(IDVentana), 10, "I", 0) & CompletarCadena(CStr(Amigo), 16, "D", " ") & RichTemporalX.TextRTF
 EnviarPaqueteTCP "43" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
                  ComandoAdicional

 ' **************************************************************
 ' Enviar el Mensaje de Multichat...
 ' **************************************************************
 'EnviarPaqueteTCP ("40" & CompletarCadena(Amigo, 16, "D", " ") & _
                       "U" & _
                       "11" & CStr(Varios.NuevoID) & RichTemporalX.TextRTF)

     
End Sub

