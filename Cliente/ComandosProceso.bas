Attribute VB_Name = "ComandosProceso"
Option Explicit
Sub ProcesarEstadoDeUnAmigo(Paquete As String)
Dim Comando, Resto, Usuario, EstadoNumero, Estadotexto, Sexo As String
Dim Respuesta, Contador As Integer
Dim Cambio As Integer

' Mensaje Recibido del Cliente:
'                       110 : Usuario No Existe
'                       111 : Ok con Info
'   Despues del 111 se envia los datos de los usuarios:
'           16 Alias Usuario
'           1  Estado
'           20 Estado Texto
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Paquete) < 2 Then
  Exit Sub
 End If
 
 Comando = Mid$(Paquete, 1, 1)
 Resto = Mid$(Paquete, 2)
 
 ' **************************************************************
 ' Comando de Usuario Inexistente
 ' **************************************************************
 Cambio = 0
 If Comando = 0 Then
  Usuario = Trim(Resto)
  For Contador = 1 To Variables.CantidadGrupoAmigo
   If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(Usuario)) Then
    If GrupoAmigo(Contador).Existe = True Then
      '
      Cambio = Contador
     Else
      '
      Cambio = -1
    End If
    GrupoAmigo(Contador).Existe = False
    Exit For
   End If
  Next
  
  If Cambio <> 0 Then ' Existe en el Listado...
   CargarAmigos ' Carga el Listado ya que Cambio....
   Exit Sub
  End If
  
  If Cambio = 0 Then ' Como no existe en el Listado de Amigos, recorre todas las ventanas
   ' Busca Todas la Ventanas y cambia y setea en todas las ventanas....
   For Contador = 0 To Forms.Count - 1
    If Forms(Contador).FormularioNombre = "Mensajes" Then
     If Forms(Contador).TratarAmigosEnChat("Buscar", CStr(Usuario)) <> 0 Then
      Forms(Contador).CargaEstadoIndividual CStr(Usuario), CStr("-1"), CStr(""), CStr("D")
      ' Verifica si hubo algun Cambio...
      If Forms(Contador).EstadoAnteriorAmigoNumero = -10 Then
        Forms(Contador).EstadoAnteriorAmigoNumero = -1
        Forms(Contador).EstadoAnteriorAmigoTexto = ""
       Else
        If Forms(Contador).EstadoAnteriorAmigoNumero <> -1 Or Forms(Contador).EstadoAnteriorAmigoTexto <> "" Then
         Forms(Contador).AvisarCambioDeEstado Usuario ', "-1", ""
        End If
      End If
     End If
    End If
   Next
  End If
 End If
 
 ' **************************************************************
 ' Valida el Paquete (Comando 1 Con Datos)
 ' **************************************************************
 If Len(Paquete) < 39 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Abre el Paquete
 ' **************************************************************
 Paquete = Mid$(Paquete, 2)
 Usuario = Trim(Mid$(Paquete, 1, 16))
 EstadoNumero = Trim(Mid$(Paquete, 17, 1))
 Estadotexto = Trim(Mid$(Paquete, 18, 20))
 Sexo = Trim(Mid$(Paquete, 38, 1))
 
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 ' Verifica si existe y/o si cambio algo
 ' **************************************************************
 Cambio = 0
 For Contador = 1 To Variables.CantidadGrupoAmigo
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(Usuario)) Then
   If GrupoAmigo(Contador).EstadoDelAmigoEstado <> EstadoNumero Or UCase(Trim(GrupoAmigo(Contador).EstadoDelAmigoTexto)) <> UCase(Trim(Estadotexto)) Or UCase(Trim(GrupoAmigo(Contador).Sexo)) <> UCase(Trim(Sexo)) Then
    ' Verifica si cambio el EstadoNumero o Texto...
    If GrupoAmigo(Contador).EstadoDelAmigoEstado <> EstadoNumero Or UCase(Trim(GrupoAmigo(Contador).EstadoDelAmigoTexto)) <> UCase(Trim(Estadotexto)) Then
     GrupoAmigo(Contador).EstadoDelAmigoEstado = EstadoNumero
     GrupoAmigo(Contador).EstadoDelAmigoTexto = UCase(Trim(Estadotexto))
     Cambio = Contador
    End If
    ' Verifica si cambio el Sexo...
    If UCase(Trim(GrupoAmigo(Contador).Sexo)) <> UCase(Trim(Sexo)) Then
     GrupoAmigo(Contador).Sexo = UCase(Trim(Sexo))
     Cambio = -1
    End If
    Exit For
   End If
  End If
 Next
 
 ' **************************************************************
 ' Si el Usuario Existe, Simplemente Carga el Listado de Amigos, por lo cual
 ' se genera automaticamente el Mensaje...
 ' **************************************************************
 If Cambio = Contador Or Cambio = -1 Then
  CargarAmigos ' Carga el Listado ya que Cambio....
  Exit Sub ' Sale...
 End If
  
 ' **************************************************************
 ' Aqui procesa en todos los formularios el camibio de estado
 ' del amigo, ya que el mismo no esta en el Listado de Amigos...
 ' **************************************************************
 For Contador = 0 To Forms.Count - 1
   If Forms(Contador).FormularioNombre = "Mensajes" Then
    Dim UsuarioPosicion As Integer
    UsuarioPosicion = Forms(Contador).TratarAmigosEnChat("Buscar", CStr(Usuario))
    If UsuarioPosicion <> 0 Then
     Dim EstadoAnteriorTexto, EstadoAnteriorNumero As String
     EstadoAnteriorNumero = Forms(Contador).BuscarEstadoNumerooTexto(CStr(Usuario), "NUMERO")
     EstadoAnteriorTexto = Forms(Contador).BuscarEstadoNumerooTexto(CStr(Usuario), "TEXTO")
     ' Verifica que haya cambiado el Estado...
     Cambio = False
     If EstadoAnteriorNumero <> "-1" Then
      If CInt(EstadoAnteriorNumero) <> CInt(EstadoNumero) Then Cambio = True
      If UCase(Trim(EstadoAnteriorTexto)) <> UCase(Trim(Estadotexto)) And CInt(EstadoNumero) = 3 Then Cambio = True
     End If
          
     ' Pone el Nuevo Estado...
     Forms(Contador).CargaEstadoIndividual CStr(Usuario), CStr(EstadoNumero), CStr(Estadotexto), CStr(Sexo)
          
     ' Cambio !! y lo informa...
     If Cambio Then
      Forms(Contador).AvisarCambioDeEstado CStr(Usuario) ', GrupoAmigo(Cambio).EstadoDelAmigoEstado, GrupoAmigo(Cambio).EstadoDelAmigoTexto
     End If
          
     ' --EstadoNumerico As String
     ' --Sexo As String
     ' --Estadotexto As String
     
     
     '' Verifica si hubo algun Cambio...
     ''Debug.Print Forms(Contador).EstadoAnteriorAmigoNumero
     'If Forms(Contador).EstadoAnteriorAmigoNumero = -10 Then
     '  Forms(Contador).EstadoAnteriorAmigoNumero = EstadoNumero
     '  Forms(Contador).EstadoAnteriorAmigoTexto = Estadotexto
     ' Else
     '  Debug.Print Forms(Contador).EstadoAnteriorAmigoNumero
     '  Debug.Print EstadoNumero
     '  Debug.Print Forms(Contador).EstadoAnteriorAmigoTexto
     '  Debug.Print Estadotexto
     '  If CInt(Forms(Contador).EstadoAnteriorAmigoNumero) <> CInt(EstadoNumero) Or UCase(Trim(Forms(Contador).EstadoAnteriorAmigoTexto)) <> UCase(Trim(Estadotexto)) Then
     '   Forms(Contador).AvisarCambioDeEstado CStr(Usuario), GrupoAmigo(Cambio).EstadoDelAmigoEstado, GrupoAmigo(Cambio).EstadoDelAmigoTexto
     '   Forms(Contador).EstadoAnteriorAmigoNumero = EstadoNumero
     '   Forms(Contador).EstadoAnteriorAmigoTexto = Estadotexto
     '  End If
     'End If
    End If
   End If
 Next
 
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 
End Sub
Sub ProcesarBusquedaAmigos(Paquete As String)
Dim TamanioTMP  As Long
Dim Tamanio, Contador, VentanaID As Integer
Dim BusquedaTMP As String
Dim Estado, Nombre, Alias, Texto As String

' Mensaje Recibido del Cliente:
'                       250 : No Hay Coincidencias
'                       251 : Hay Coincidencias
'   Despues del 251 se envia los datos de los usuarios:
'           16 Alias Usuario
'           1  Estado
'           50 Apellido y Nombre

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Paquete) < 1 Then
  ' Descarta el Paquete
  Exit Sub
 End If
 If Not IsNumeric(Mid$(Paquete, 1, 1)) Then
  ' Descarta el Paquete
  Exit Sub
 End If
  
 ' **************************************************************
 ' Pone el Resultado de la Busqueda
 ' **************************************************************
 Variables.RespuestaBusquedaAmigos = CInt(Mid$(Paquete, 1, 1))
 ' No hay usuarios Encontrados
 If CInt(Mid$(Paquete, 1, 1)) = 0 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica el Formato del Paquete
 ' **************************************************************
 Paquete = Mid$(Paquete, 2)
 Tamanio = Len(Paquete)
 TamanioTMP = Tamanio / 67
 If TamanioTMP <> Int(Tamanio / 67) Then
  ' El Paquete es Invalido se descarta...
  Exit Sub
 End If
 
 ' **************************************************************
 ' Ubica la Ventana Donde tiene que descargar la Busqueda
 ' **************************************************************
 VentanaID = -1
 For Contador = 1 To Forms.Count
  If Trim(UCase(Forms(Contador).FormularioNombre)) = UCase("AgregarBuscarAmigos") Then
   VentanaID = Contador
   Exit For
  End If
 Next
 ' Verifica que se haya encontrado la Ventana
 If VentanaID = -1 Then Exit Sub
 
 ' **************************************************************
 ' Descarga el Listado...
 ' **************************************************************
 ' Borra la Lista
 Forms(VentanaID).ResultadoBusqueda.Clear
 ' Craga la Lista
 For Contador = 0 To TamanioTMP - 1
  BusquedaTMP = Mid$(Paquete, (Contador * 67) + 1, 67)
  ' **************************************************************
  ' Define el Texto a Mostra
  ' **************************************************************
  Alias = Trim(Mid$(BusquedaTMP, 1, 16))
  Estado = Trim(Mid$(BusquedaTMP, 17, 1))
  Nombre = Trim(Mid$(BusquedaTMP, 18))
  Texto = "[" & Alias & "] "
  If Estado = "0" Then
    ' (Desconectado)
    Texto = Texto & MensajeRecurso(308)
   Else
    Texto = Texto & MensajeRecurso(309)
  End If
  Texto = Texto & " - '" & Nombre & "'..."
  ' **************************************************************
  Forms(VentanaID).ResultadoBusqueda.AddItem Texto
 Next
 
End Sub
Sub ProcesarValidarAmigo(Paquete As String)
Dim Respuesta As Integer
Dim Resultado, Estado, Estadotexto, Sexo
Dim Alias As String

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Paquete) <> 39 Then
  ' Descarta el Paquete
  Exit Sub
 End If
 
 ' **************************************************************
 ' Procesa el Resultado
 ' **************************************************************
 Resultado = Mid$(Paquete, 1, 1)
 Estado = Mid$(Paquete, 2, 1)
 Estadotexto = Trim(Mid$(Paquete, 3, 20))
 Sexo = Trim(Mid$(Paquete, 23, 1))
 Alias = Trim(Mid$(Paquete, 24, 16))
 
 ' **************************************************************
 ' Verifica que Resulado y Estado Sean Numericos
 ' **************************************************************
 If Not IsNumeric(Resultado) Or Not IsNumeric(Estado) Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Carga el Resultado
 ' **************************************************************
 Variables.RespuestaAmigoACrearAlias = Trim(Alias)
 Variables.RespuestaAmigoACrearResultado = CInt(Resultado)
 Variables.RespuestaAmigoACrearEstado = CInt(Estado)
 Variables.RespuestaAmigoACrearEstadoTexto = Estadotexto
 Variables.RespuestaAmigoACrearSexo = Sexo
 
End Sub
Sub ProcesarPeticionDeListadoDeAmigos(Paquete As String)
Dim Tamanio, Contador, Contador3 As Integer
Dim TamanioTMP  As Long
Dim PosicionTMP As Integer
Dim GrupoAmigoTemp() As Variables.MiGrupoAmigo
Dim CantidadGrupoAmigoTemp As Integer
Dim ContadorTMP, Actual As Integer

' Formato del Paquete:
'   NombreDelGrupo 20
'   IDNombreDelAmigo 16
'   EstadoDelAmigoEstado 1
'   EstadoDelAmigoTexto 20
'   NombreDelAmigo 50
'   Sexo 1
'   Existe 1
'   Direccion Email 50
'   Total del Paquete 159

 ' **************************************************************
 ' Define los Amigos Anteriores
 ' **************************************************************
 ReDim GrupoAmigoTMP(Variables.CantidadGrupoAmigo)
 GrupoAmigoTemp() = Variables.GrupoAmigo()
 CantidadGrupoAmigoTemp = Variables.CantidadGrupoAmigo
 
 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 Tamanio = Len(Paquete)
 TamanioTMP = Tamanio / 159
 If TamanioTMP <> Int(Tamanio / 159) Then
  ' El Paquete es Invalido se descarta...
  Exit Sub
 End If
 
 PosicionTMP = 0
 ' Define la Nueva cantidad de Amigos
 Variables.CantidadGrupoAmigo = TamanioTMP
 ReDim GrupoAmigo(Variables.CantidadGrupoAmigo)
 ' Comienza el Proceso de los Amigos
 For Contador = 0 To TamanioTMP - 1
  PosicionTMP = Contador * 159 + 1
  ' **************************************************************
  ' Carga los datos de los Amigos
  ' **************************************************************
  ' Nombre del Grupo
  Variables.GrupoAmigo(Contador + 1).NombreDelGrupo = Mid$(Paquete, PosicionTMP, 20)
  ' ID Nombre del AMigo
  Variables.GrupoAmigo(Contador + 1).IDNombreDelAmigo = Mid$(Paquete, PosicionTMP + 20, 16)
  ' Estado del Amigo
  Variables.GrupoAmigo(Contador + 1).EstadoDelAmigoEstado = CInt(Mid$(Paquete, PosicionTMP + 20 + 16, 1))
  ' Estado del Amigo Texo
  Variables.GrupoAmigo(Contador + 1).EstadoDelAmigoTexto = Mid$(Paquete, PosicionTMP + 20 + 16 + 1, 20)
  ' Nombre del Amigo
  Variables.GrupoAmigo(Contador + 1).NombreDelAmigo = Mid$(Paquete, PosicionTMP + 20 + 16 + 1 + 20, 50)
  ' Nombre Sexo
  Variables.GrupoAmigo(Contador + 1).Sexo = Mid$(Paquete, PosicionTMP + 20 + 16 + 1 + 20 + 50, 1)
  ' Nombre Existe
  If Mid$(Paquete, PosicionTMP + 20 + 16 + 1 + 20 + 50 + 1, 1) = 1 Then
    Variables.GrupoAmigo(Contador + 1).Existe = True
   Else
    Variables.GrupoAmigo(Contador + 1).Existe = False
  End If
  ' Direccion de Email
  Variables.GrupoAmigo(Contador + 1).DireccionEMail = Mid$(Paquete, PosicionTMP + 20 + 16 + 1 + 20 + 50 + 2)
 Next
  
 ' **************************************************************
 ' Carga los Grupos actuales con el estado (Expandido o no)
 ' **************************************************************
 Varios.CargarEstadoDeNodos
 
 ' **************************************************************
 ' Carga los Amigos
 ' **************************************************************
 Varios.CargarAmigos
 
 ' **************************************************************
 ' Pone los Mensaje de Coneccion/Desconeccion de los Amigos
 ' **************************************************************
 ' Verifica que no sea la Primera Carga
 If Cliente.ListadoDeAmigos.Nodes.Count = 0 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Pone los Mensaje de Coneccion/Desconeccion de los Amigos
 ' **************************************************************
 For Contador = 1 To Variables.CantidadGrupoAmigo
  For ContadorTMP = 1 To CantidadGrupoAmigoTemp
   ' Busca el Amigo de GrupoAmigo en GrupoAmigoTemo
   If Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo) = Trim(GrupoAmigoTemp(ContadorTMP).IDNombreDelAmigo) Then
    ' Verifica si cambio el Estado
    If (Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado) <> Trim(GrupoAmigoTemp(ContadorTMP).EstadoDelAmigoEstado)) Or (Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoTexto) <> Trim(GrupoAmigoTemp(ContadorTMP).EstadoDelAmigoTexto)) Then
     ' Crea y el Nuevo Formulario
     Dim MostrarTiras As New InformarCambioDeEstado
     Load MostrarTiras
     ' Define el Form que se creo
     Actual = Forms.Count - 1
     ' El Estado se Cambio: Muestra la Tira
     ' **************************************************************
     ' Verifica si se quiere se informe de los Cambios de Estado
     ' **************************************************************
     If Configuracion.InformarCambiosDeEstado = True Then
      Forms(Actual).MostrarFormulario Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo), Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado), Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoTexto), Trim(Variables.GrupoAmigo(Contador).Sexo)
     End If
     ' **************************************************************
     ' Busca en todas las ventanas de Mensajes para avisar del cambio...
     ' **************************************************************
     For Contador3 = 0 To Forms.Count - 1
      If Forms(Contador3).FormularioNombre = "Mensajes" Then
       If Forms(Contador3).TratarAmigosEnChat("Buscar", Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo)) <> 0 Then
        ' Ok Esta... Debe Cambiarlo...
        Forms(Contador3).AvisarCambioDeEstado Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo) ', Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado, Variables.GrupoAmigo(Contador).EstadoDelAmigoTexto
       End If
      End If
     Next
     ' Sale del Bucle de Busqueda Anidado...
     Exit For
    End If
   End If
  Next
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
Sub ProcesarGrabacionDeListadoDeAmigos(Paquete As String)
Dim Comando As String
 
 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Paquete) <> 1 Then
  ' El largo del Paquete es incorrecto
  Exit Sub
 End If
 
 ' **************************************************************
 ' Procesar
 ' **************************************************************
 Comando = Mid$(Paquete, 1, 1)
 Select Case Comando
  Case "1":
   ' ok Todo Bien...
  Case Else
   ' No Fue Posible Realizar la Grabacion de los Cambios en su Listado de Amigos...
   MostrarMSGBox MensajeRecurso(310), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
 End Select

End Sub
Sub ProcesarDatosUsuario(Estado As String, Paquete As String)
Dim Contador As Integer
Dim Formulario As Integer

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 If Len(Paquete) < 16 Then
  ' El largo del Paquete es incorrecto
  Exit Sub
 End If
 
 ' **************************************************************
 ' Busca el Formulario donde poner los datos
 ' **************************************************************
 Formulario = 0
 For Contador = 1 To Forms.Count - 1
  Dim FormularioNombre, AliasUsuario, UsuarioAliasPaquete As String
  ' **************************************************************
  ' Verifica que el Formulario sea de Datos
  ' **************************************************************
  If Forms(Contador).FormularioNombre = "DatosUsuario" Then
   FormularioNombre = Trim(Forms(Contador).FormularioNombre)
   AliasUsuario = Trim(Forms(Contador).AliasUsuario)
   UsuarioAliasPaquete = Trim(Mid$(Paquete, 1, 16))
   If UCase(FormularioNombre) = UCase("DatosUsuario") And UCase(AliasUsuario) = UCase(UsuarioAliasPaquete) Then
    ProcesaLosDatosEnElFormulario Contador, Estado, UsuarioAliasPaquete, Paquete
    Formulario = 1
   End If
  End If
 Next
 
 ' Si el Formulario no existe, Descarta la info
 
End Sub
Sub ProcesarGrabacionUsuario(Estado As String, PaqueteRecibido As String)
Dim Contador As Integer
Dim Paquete As String


 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 Paquete = PaqueteRecibido
 If Len(Paquete) <> 1 Then
  ' El largo del Paquete es incorrecto
  Exit Sub
 End If
 
 ' **************************************************************
 ' Busca el Formulario donde Avisar que se cambiaron los Datos
 ' **************************************************************
 For Contador = 1 To Forms.Count - 1
  ' **************************************************************
  ' Verifica que el Formulario sea de Cambio de Datos
  ' **************************************************************
  If Forms(Contador).FormularioNombre = "DatosUsuario" Then
   If Forms(Contador).CambioDeDatosUsuario = True Then
    ' Confirma la Grabacion de la Info
    If Paquete = "1" Then
     Forms(Contador).GraboLosDatosUsuario = True
    End If
    ' Si hubo un error deja que salga por Time-Out
   End If
  End If
 Next
 
End Sub
Sub ProcesaLosDatosEnElFormulario(Formulario As Integer, Estado As String, UsuarioAliasPaquete As String, PaqueteRecibido As String)
Dim Paquete As String
Dim Contador As Integer

 Paquete = Mid$(PaqueteRecibido, 17)
 
 ' **************************************************************
 ' Procesa el Paquete - Define que el Usuario No existe
 ' **************************************************************
 ' El Usuario no Existe
 If Estado = "2" Then
  ' **************************************************************
  ' Detiene los Controles del Formulario
  ' **************************************************************
  Forms(Formulario).TimeOut.Enabled = False
  Forms(Formulario).Animacion.Enabled = False
  Forms(Formulario).Refrescando = False
  ' Pone la Imagen de Desconectado
  Forms(Formulario).AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
  ' El Usuario %  No Existe...
  MostrarMSGBox MensajeRecurso(174) & Trim(UsuarioAliasPaquete) & MensajeRecurso(312), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Forms(Formulario).BlanquearCampos
  Forms(Formulario).EstadoUsuarioImagen = Cliente.ImagenesAmigos.ListImages("UsuarioNoExiste").Picture
  Forms(Formulario).EstadoUsuarioTexto = MensajeRecurso(342)
  ' Verifica si cambio el Estado en el Listado, si es asi, comom tiene lainfo lacambia...
  For Contador = 1 To Variables.CantidadGrupoAmigo
   If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(UsuarioAliasPaquete)) Then
    If GrupoAmigo(Contador).Existe <> False Then
     GrupoAmigo(Contador).Existe = False
     CargarAmigos
     Exit For
    End If
   End If
  Next
  Exit Sub
 End If
 
 ' Pone los Datos en el Usuario
 If Estado = "1" Then
  Forms(Formulario).ApellidoYNombre = Trim(Mid$(Paquete, 1, 50))
  Forms(Formulario).DireccionDeEmail = Trim(Mid$(Paquete, 51, 50))
  Forms(Formulario).Edad = Trim(Mid$(Paquete, 101, 2))
  If Trim(Mid$(Paquete, 103, 1)) = "F" Then
    Forms(Formulario).Sexo = MensajeRecurso(249)
    If UCase(Trim(Forms(Formulario).IDAliasUsuario)) = UCase(Trim(Configuracion.IDAliasUsuario)) Then
     Cliente.ListadoDeAmigos.Nodes(2).Image = "Mujer"
     Cliente.ListadoDeAmigos.Nodes(2).SelectedImage = "Mujer"
    End If
   Else
    Forms(Formulario).Sexo = MensajeRecurso(248)
    If UCase(Trim(Forms(Formulario).IDAliasUsuario)) = UCase(Trim(Configuracion.IDAliasUsuario)) Then
     Cliente.ListadoDeAmigos.Nodes(2).Image = "Hombre"
     Cliente.ListadoDeAmigos.Nodes(2).SelectedImage = "Hombre"
    End If
  End If
  Forms(Formulario).UbicacionGeografica = Trim(Mid$(Paquete, 104, 20))
  Forms(Formulario).Intencion = Trim(Mid$(Paquete, 124, 20))
  Forms(Formulario).Humor = Trim(Mid$(Paquete, 144, 20))
  Forms(Formulario).Ocupacion = Trim(Mid$(Paquete, 164, 20))
  
  ' **************************************************************
  ' Arregla el Signo del Usuario...
  ' **************************************************************
  Dim Signo As String
  'Dim Contador As Integer
  Signo = Trim(Mid$(Paquete, 184, 15))
  For Contador = 0 To 11
   If UCase(Trim(MensajeRecursoReal(254 + Contador))) = UCase(Trim(Signo)) Then
    Signo = MensajeRecurso(Contador + 254)
   End If
  Next
  Forms(Formulario).Signo = Signo
  
  Select Case Trim(Mid$(Paquete, 199, 1))
   Case "C"
    Forms(Formulario).EstadoCivil = MensajeRecurso(250)
   Case "D"
    Forms(Formulario).EstadoCivil = MensajeRecurso(251)
   Case "V"
    Forms(Formulario).EstadoCivil = MensajeRecurso(253)
   Case "S"
    Forms(Formulario).EstadoCivil = MensajeRecurso(252)
  End Select
  Forms(Formulario).Telefono = Trim(Mid$(Paquete, 200, 50))
  Forms(Formulario).OtraInfo = Trim(Mid$(Paquete, 250, 150))
  Forms(Formulario).FechaDeNacimiento = Trim(Mid$(Paquete, 400, 10))
  Forms(Formulario).EstadoNumero = Trim(Mid$(Paquete, 410, 1))
  Dim EstadoUsuarioDato As String
  ' Pone el Estado en Imagen
  EstadoUsuarioDato = Trim(Mid$(Paquete, 410, 1))
  ' Verifica que exista un Sexo Correcto
  If Trim(Mid$(Paquete, 103, 1)) <> "F" And Trim(Mid$(Paquete, 103, 1)) <> "M" Then
    If EstadoUsuarioDato = "0" Then Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("NoConectado").Picture
    If EstadoUsuarioDato = "1" Then Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoVisible").Picture
    If EstadoUsuarioDato = "2" Then Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoNoDisponible").Picture
    If EstadoUsuarioDato = "3" Then Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoCustom").Picture
   Else
    Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages(Trim(Mid$(Paquete, 103, 1)) & EstadoUsuarioDato).Picture
  End If
  Select Case EstadoUsuarioDato
   Case "0":
    'Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("NoConectado").Picture
    ' No Conectado...
    Forms(Formulario).EstadoUsuarioTexto = " - " & MensajeRecurso(287)
   Case "1":
    'Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoVisible").Picture
    ' Disponible (Normal)...
    Forms(Formulario).EstadoUsuarioTexto = " - " & MensajeRecurso(180)
   Case "2"
    'Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoNoDisponible").Picture
    ' No Disponible..
    Forms(Formulario).EstadoUsuarioTexto = " - " & MensajeRecurso(181)
   Case "3"
    'Forms(Formulario).EstadoUsuarioImagen.Picture = Cliente.Imagenes.ListImages("EstadoCustom").Picture
    ' Custom...
    Forms(Formulario).EstadoUsuarioTexto = " - " & ArreglarLenguaje(Trim(Mid$(Paquete, 411, 20)))
  End Select
 End If
 
 ' **************************************************************
 ' Si Cambio Hace el Cambio Respectivo...
 ' **************************************************************
 Dim Cambio As Boolean
 For Contador = 1 To Variables.CantidadGrupoAmigo
  Cambio = False
  If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(UsuarioAliasPaquete)) Then
   ' Direccion de EMAIL...
   If UCase(Trim(GrupoAmigo(Contador).DireccionEMail)) <> UCase(Trim(Mid$(Paquete, 51, 50))) Then
    GrupoAmigo(Contador).DireccionEMail = Trim(Mid$(Paquete, 51, 50))
    Cambio = True
   End If
   ' Estado Numero...
   If CInt(GrupoAmigo(Contador).EstadoDelAmigoEstado) <> CInt(Trim(Mid$(Paquete, 410, 1))) Then
    GrupoAmigo(Contador).EstadoDelAmigoEstado = Trim(Mid$(Paquete, 410, 1))
    Cambio = True
   End If
   ' Estado Texto...
   If UCase(Trim(GrupoAmigo(Contador).EstadoDelAmigoTexto)) <> Mid$(UCase(Trim(Forms(Formulario).EstadoUsuarioTexto)), 3) Then
    GrupoAmigo(Contador).EstadoDelAmigoTexto = Mid$(Trim(Forms(Formulario).EstadoUsuarioTexto), 3)
    Cambio = True
   End If
   ' Sexo
   If UCase(Trim(GrupoAmigo(Contador).Sexo)) <> UCase(Trim(Mid$(Paquete, 103, 1))) Then
    GrupoAmigo(Contador).Sexo = Trim(Mid$(Paquete, 103, 1))
    Cambio = True
   End If
   ' Existe?
   If GrupoAmigo(Contador).Existe <> True Then
    GrupoAmigo(Contador).Existe = True
    Cambio = True
   End If
   ' Nombre y Apellido
   If UCase(Trim(GrupoAmigo(Contador).NombreDelAmigo)) <> UCase(Trim(Mid$(Paquete, 1, 50))) Then
    GrupoAmigo(Contador).NombreDelAmigo = Trim(Mid$(Paquete, 1, 50))
    Cambio = True
   End If
   If Cambio Then
    CargarAmigos
    Exit For
   End If
  End If
 Next
  
 ' **************************************************************
 ' Detiene los Controles del Formulario
 ' **************************************************************
 Forms(Formulario).TimeOut.Enabled = False
 Forms(Formulario).Animacion.Enabled = False
 Forms(Formulario).Refrescando = False
 ' Pone la Imagen de Desconectado
 Forms(Formulario).AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture

 
End Sub
Sub ConfirmarElEnvioDeMail(Paquete As String)
Dim Comando, Datos As String

 ' Comandos recibidos...
 '       0: No se pudo enviar la Password...
 '       1: La password fue enviada a enviopassworddireccionmail
 '       2: El usuario no posee direccion de Email
 '       3: EL usuario no existe...
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Paquete) < 1 Then Exit Sub
 
 ' **************************************************************
 ' Abre el Paquete
 ' **************************************************************
 Comando = Mid$(Paquete, 1, 1)
 Datos = Mid$(Paquete, 2)
 
 ' **************************************************************
 ' Procesa los Comandos de Error
 ' **************************************************************
 Select Case Comando
  Case "3" Or "2":
   Loguin.EnvioPassword = CInt(Comando)
   Loguin.EnvioPasswordDireccionMail = ""
   Exit Sub
 End Select
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Datos) <> 50 Then Exit Sub
 
 ' **************************************************************
 ' Procesa los Comandos de OK - o Error con Direccion de Mail
 ' **************************************************************
 Select Case Comando
  Case "0" Or "1":
   Loguin.EnvioPassword = CInt(Comando)
   Loguin.EnvioPasswordDireccionMail = Trim(Datos)
   Exit Sub
 End Select
  
End Sub

