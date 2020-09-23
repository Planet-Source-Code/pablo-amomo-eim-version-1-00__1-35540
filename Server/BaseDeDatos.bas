Attribute VB_Name = "BaseDeDatos"
Option Explicit
Function EnviaMensajesOffline(UsuarioID As String, Port As Integer)
Dim rsMensajeOffLineTMP As ADODB.Recordset
Dim StringSQL As String

  ' **************************************************************
  ' Crea el Recorset de los Mensajes Offline...
  ' **************************************************************
  Set rsMensajeOffLineTMP = New ADODB.Recordset
  rsMensajeOffLineTMP.CursorType = adOpenKeyset
  rsMensajeOffLineTMP.LockType = adLockOptimistic
  StringSQL = "SELECT * FROM MensajesOffLine Where [UsuarioPara]='" & UCase(Trim(UsuarioID)) & "' order by [FechaYHora]"
  rsMensajeOffLineTMP.Open StringSQL, DataBase
  ' **************************************************************
  ' Refresca los MensajesOffLine...
  ' **************************************************************
  If rsMensajeOffLineTMP.State = adStateOpen Then
   rsMensajeOffLineTMP.Requery
  End If
  
  ' **************************************************************
  ' Si hay Mensaje los Manda....
  ' **************************************************************
  If rsMensajeOffLineTMP.RecordCount <> 0 Then
   rsMensajeOffLineTMP.MoveFirst
   Do Until rsMensajeOffLineTMP.EOF
    Dim tiempoinicial As Date
    tiempoinicial = Time
    Do Until DateDiff("s", tiempoinicial, Time) > 1
     DoEvents
    Loop
    Dim Usuario, FechaYHora, Mensaje As String
    Usuario = Trim(rsMensajeOffLineTMP![UsuarioEmisor])
    FechaYHora = Trim(rsMensajeOffLineTMP![FechaYHora])
    Mensaje = Trim(rsMensajeOffLineTMP![Mensaje])
    ' **************************************************************
    ' Envia el Mensaje Offline...
    ' **************************************************************
    Dim UsuarioEmisor As Integer
    EnviarPaqueteTCP "32" & CompletarCadena("EIM", 16, "D", " ") & "52" & CompletarCadena(CStr(Usuario), 16, "D", " ") & CompletarCadena(CStr(FechaYHora), 19, "D", " ") & Mensaje, Port
    ' **************************************************************
    ' Espera un Segundo cuando el Emisor y el Receptor es el Mismo
    ' **************************************************************
    rsMensajeOffLineTMP.MoveNext
   Loop
  End If
  
  ' **************************************************************
  ' Borra los Mensaje enviados...
  ' **************************************************************
  Set rsMensajeOffLineTMP = New ADODB.Recordset
  rsMensajeOffLineTMP.CursorType = adOpenKeyset
  rsMensajeOffLineTMP.LockType = adLockOptimistic
  StringSQL = "Delete * FROM MensajesOffLine Where [UsuarioPara]='" & UCase(Trim(UsuarioID)) & "'"
  rsMensajeOffLineTMP.Open StringSQL, DataBase
  'rsMensajeOffLineTMP.Close
  
End Function
Function GrabarListadoDeAmigos(IDUsuarioNumerico As Integer, Listado As String) As Integer
 
 ' **************************************************************
 ' Pararse en el Recordset Correspondiente
 ' **************************************************************
 ' Se posiciona al Principio de los Registros
 rsTablaUsuarios.MoveFirst
 ' Se Posiciona x lugares empezando del Principio
 rsTablaUsuarios.Move IDUsuarioNumerico - 1
 
 ' **************************************************************
 ' Graba el Nuevo Listado de Amigos
 ' **************************************************************
 With Usuarios(IDUsuarioNumerico)
  rsTablaUsuarios![ListadoDeAmigos] = .ListadoDeAmigos
 End With
 rsTablaUsuarios.Update
 
 ' **************************************************************
 ' Informa que la Grabacion Fue Ok
 ' **************************************************************
 GrabarListadoDeAmigos = 1
 
End Function
Function GrabarFechaUltimoLogueo(IDUsuarioNumerico As Integer) As Integer
' Devuelve:
'           1 : Todo OK
'           0 : No Se Grabo
 
 ' **************************************************************
 ' Pararse en el Recordset Correspondiente
 ' **************************************************************
 ' Se posiciona al Principio de los Registros
 rsTablaUsuarios.MoveFirst
 ' Se Posiciona x lugares empezando del Principio
 rsTablaUsuarios.Move IDUsuarioNumerico - 1

 ' **************************************************************
 ' Realiza el update de la Fecha de Logueo
 ' **************************************************************
 With Usuarios(IDUsuarioNumerico)
  rsTablaUsuarios![UltimoLogueo] = Varios.FechaActual
 End With
 rsTablaUsuarios.Update
 
 ' **************************************************************
 ' Devuelve Todo OK...
 ' **************************************************************
 GrabarFechaUltimoLogueo = 1
 
End Function
Function GrabarModificacionesUsuario(IDUsuarioNumerico As Integer, Optional Borrando As Boolean) As Integer
' Devuelve:
'           1 : Todo OK
'           0 : No Se Grabo

 ' **************************************************************
 ' Pararse en el Recordset Correspondiente
 ' **************************************************************
 ' Se posiciona al Principio de los Registros
 rsTablaUsuarios.MoveFirst
 ' Se Posiciona x lugares empezando del Principio
 rsTablaUsuarios.Move IDUsuarioNumerico - 1

 ' **************************************************************
 ' Realiza el update del Password
 ' **************************************************************
 With Usuarios(IDUsuarioNumerico)
  rsTablaUsuarios![ApellidoYNombre] = .ApellidoYNombre
  rsTablaUsuarios![DireccionDeEmail] = .DireccionDeEmail
  rsTablaUsuarios![FechaDeNacimiento] = .FechaDeNacimiento
  rsTablaUsuarios![Edad] = .Edad
  rsTablaUsuarios![EstadoCivil] = .EstadoCivil
  rsTablaUsuarios![Humor] = .Humor
  rsTablaUsuarios![IDAliasUsuario] = .IDAliasUsuario
  rsTablaUsuarios![Intencion] = .Intencion
  rsTablaUsuarios![Ocupacion] = .Ocupacion
  rsTablaUsuarios![OtraInfo] = .OtraInfo
  rsTablaUsuarios![Sexo] = .Sexo
  rsTablaUsuarios![SigNo] = .SigNo
  rsTablaUsuarios![Telefono] = .Telefono
  rsTablaUsuarios![UbicacionGeografica] = .UbicacionGeografica
  rsTablaUsuarios![ListadoDeAmigos] = .ListadoDeAmigos
  rsTablaUsuarios![Password] = .Password
  rsTablaUsuarios![MensajesOffline] = .MensajesOffline
  rsTablaUsuarios![UsuarioBloqueado] = .UsuarioBloqueado
  rsTablaUsuarios![Password] = .Password
  If Borrando = True Then
   rsTablaUsuarios![UltimoLogueo] = .UltimoLogueo
  End If
 End With
 rsTablaUsuarios.Update
 
 GrabarModificacionesUsuario = 1
 
End Function
Function GrabarNuevaPassword(IDUsuarioNumerico As Integer, NuevaPassword As String) As String
' Devuelve:
'           1 : Todo OK
'           0 : No Se Grabo


 ' **************************************************************
 ' Pararse en el Recordset Correspondiente
 ' **************************************************************
 ' Se posiciona al Principio de los Registros
 rsTablaUsuarios.MoveFirst
 ' Se Posiciona x lugares empezando del Principio
 rsTablaUsuarios.Move IDUsuarioNumerico - 1

 ' **************************************************************
 ' Realiza el update del Password
 ' **************************************************************
 rsTablaUsuarios![Password] = NuevaPassword
 rsTablaUsuarios.Update
 Usuarios(IDUsuarioNumerico).Password = NuevaPassword
 
 ' **************************************************************
 ' Devuelve un OK en la Grabación
 ' **************************************************************
 GrabarNuevaPassword = 1
 
End Function
Sub BajarBaseDeDatos()

 ' **************************************************************
 ' Cierra las Instancias de la Base de Datos
 ' **************************************************************
 If DataBase.State = adStateClosed And rsTablaUsuarios.State = adStateClosed Then
  Exit Sub ' La Base ya esta cerrada...
 End If
 
 ' **************************************************************
 ' Cierra las Instancias de la Base de Datos
 ' **************************************************************
 rsTablaUsuarios.Close
 DataBase.Close
  
End Sub
Function AbrirBaseDeDatos() As Boolean
'On Error GoTo AbrirBasedDatos_Error
Dim StringBase, ConeccionString, Temp As String
Dim Existe As Boolean
 
 ' **************************************************************
 ' Carga la Ubicacion de la Base de Datos...
 ' **************************************************************
 StringBase = Trim(Configuracion.UbicacionBaseDeDatos) & Trim(Configuracion.NombreDeLaBaseDeDatos)
 
 ' **************************************************************
 ' Verifica que la Base de Datos exista en la Ubicacion Actual
 ' **************************************************************
 Temp = Dir(Trim(Configuracion.UbicacionBaseDeDatos) & Trim(Configuracion.NombreDeLaBaseDeDatos))
 If UCase(Temp) <> UCase(Trim(Configuracion.NombreDeLaBaseDeDatos)) Then
  ' El Archivo no Existe en la Direccion definida por la Configuracion...
  Temp = Dir(App.Path & "\Database\" & Trim(Configuracion.NombreDeLaBaseDeDatos))
  If UCase(Temp) <> UCase(Trim(Configuracion.NombreDeLaBaseDeDatos)) Then
    ' El Archivo no existe en el directorio de la Aplicacion...
    MsgBox "The Database File Don't Exist... Please Review the Configuration...(The System Can't Start...)", vbCritical, Configuracion.TituloVentanas
    Logs.EscribirEvento "The Database File Don't Exist... Please Review the Configuration...(The System Can't Start...)", vbRed
    AbrirBaseDeDatos = False
    Exit Function
   Else
    ' Arranca el Sistema desde la App.Path...
    Configuracion.UbicacionBaseDeDatos = App.Path & "\Database\"
    ConfiguracionEnArchivo.GrabarConfiguracionArchivo
    StringBase = Trim(Configuracion.UbicacionBaseDeDatos) & Trim(Configuracion.NombreDeLaBaseDeDatos)
  End If
 End If
 
 ' **************************************************************
 ' Definir el String de Conneccion con la Info requerida...
 ' **************************************************************
 ConeccionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StringBase & ";Persist Security Info=False"
 
 
 ' **************************************************************
 ' Carga los Parametros de Configuracion de la Base y Abre la
 ' Coneccion... (Define el Path de la Base de Datos...)
 ' **************************************************************
 Set DataBase = New ADODB.Connection
 DataBase.Open ConeccionString

 ' **************************************************************
 
 ' **************************************************************
 ' Abre el Recordset de Usuarios
 ' **************************************************************
 Set rsTablaUsuarios = New ADODB.Recordset
 rsTablaUsuarios.CursorType = adOpenKeyset
 rsTablaUsuarios.LockType = adLockOptimistic
 rsTablaUsuarios.Open "SELECT * FROM Usuarios ORDER BY IDNumericoUsuario", DataBase
 ' **************************************************************
 ' **************************************************************
 ' Abre el Recordset de Logs
 ' **************************************************************
 Set rsTablaLogs = New ADODB.Recordset
 rsTablaLogs.CursorType = adOpenKeyset
 rsTablaLogs.LockType = adLockOptimistic
 rsTablaLogs.Open "SELECT * FROM logs", DataBase
 ' **************************************************************
 ' **************************************************************
 ' Abre la base de Mensajes Offline...
 ' **************************************************************
 Set rsMensajeOffLine = New ADODB.Recordset
 rsMensajeOffLine.CursorType = adOpenKeyset
 rsMensajeOffLine.LockType = adLockOptimistic
 rsMensajeOffLine.Open "SELECT * FROM MensajesOffLine", DataBase
 ' **************************************************************
 
 ' **************************************************************
 ' Hace una requery - Usuarios... (Refresca Cualquier Cambio)
 ' **************************************************************
 If rsTablaUsuarios.State = adStateOpen Then
  rsTablaUsuarios.Requery
 End If '
 ' **************************************************************
 ' **************************************************************
 ' Hace una requery - Logs... (Refresca Cualquier Cambio)
 ' **************************************************************
 If rsTablaLogs.State = adStateOpen Then
  rsTablaLogs.Requery
 End If '
 ' **************************************************************
 ' **************************************************************
 ' Hace una requery - Mensajes Offline(Refresca Cualquier Cambio)
 ' **************************************************************
 If rsMensajeOffLine.State = adStateOpen Then
  rsMensajeOffLine.Requery
 End If '
 ' **************************************************************
 
 ' **************************************************************
 ' Arregla los Indices...
 ' **************************************************************
 ArreglarIndices
 
 ' **************************************************************
 ' Devuelve el Ok de la Apertura...
 ' **************************************************************
 AbrirBaseDeDatos = True
 
Salir_AbrirBasedDatos_Error:
 Exit Function

AbrirBasedDatos_Error:
 MsgBox "An Error Ocurred When Try to Open The Database...(The System Can't Start...)", vbCritical, Configuracion.TituloVentanas
 Logs.EscribirEvento "An Error Ocurred When Try to Open The Database...(The System Can't Start...)", vbRed
 AbrirBaseDeDatos = False
 Resume Salir_AbrirBasedDatos_Error

End Function
Function CargarUsuarios() As Integer
Dim CantidadDeRecords, Contador As Integer
 
 ' **************************************************************
 ' Verifica la Cantidad de Usuarios
 ' **************************************************************
 CantidadDeRecords = rsTablaUsuarios.RecordCount
 ' **************************************************************
 
 ' **************************************************************
 ' Maneja hasta un Maximo de lo definido en UsuariosSoportados...
 ' **************************************************************
 If CantidadDeRecords > CInt(Configuracion.UsuariosSoportados) Then CantidadDeRecords = CInt(Configuracion.UsuariosSoportados)
 
 ' **************************************************************
 ' Configura la Cantidad de Usuarios Actuales
 ' **************************************************************
 'Configuracion.CantidadDeUsuarios = CantidadDeRecords
 Configuracion.CantidadDeUsuarios = CInt(Configuracion.UsuariosSoportados)
 ReDim Usuarios(Configuracion.CantidadDeUsuarios)
 ' **************************************************************
 
 ' **************************************************************
 ' Limpiar la Variable de Usuario
 ' **************************************************************
 'For Contador = 1 To CantidadDeRecords
 For Contador = 1 To Configuracion.CantidadDeUsuarios
  LimpiarUsuario (Contador)
 Next
 ' **************************************************************
 
 ' **************************************************************
 ' Carga los Usuarios en la Variable Usuarios
 ' **************************************************************
 Contador = 0
 If CantidadDeRecords <> 0 Then
  rsTablaUsuarios.MoveFirst
  Do Until rsTablaUsuarios.EOF
   DoEvents
   With Usuarios(rsTablaUsuarios![IDNumericoUsuario])
    .EstadoActualNumero = 0
    .EstadoActualTexto = ""
    .PortActual = 0
    .DireccionDeEmail = NuloANada(rsTablaUsuarios![DireccionDeEmail])
    .FechaDeNacimiento = NuloANada(rsTablaUsuarios![FechaDeNacimiento])
    .Edad = NuloANada(rsTablaUsuarios![Edad])
    .EstadoCivil = NuloANada(rsTablaUsuarios![EstadoCivil])
    .Humor = NuloANada(rsTablaUsuarios![Humor])
    .IDAliasUsuario = NuloANada(rsTablaUsuarios![IDAliasUsuario])
    .IDNumericoUsuario = rsTablaUsuarios![IDNumericoUsuario]
    .Intencion = NuloANada(rsTablaUsuarios![Intencion])
    .Ocupacion = NuloANada(rsTablaUsuarios![Ocupacion])
    .OtraInfo = NuloANada(rsTablaUsuarios![OtraInfo])
    .Sexo = NuloANada(rsTablaUsuarios![Sexo])
    .SigNo = NuloANada(rsTablaUsuarios![SigNo])
    .Telefono = NuloANada(rsTablaUsuarios![Telefono])
    .UbicacionGeografica = NuloANada(rsTablaUsuarios![UbicacionGeografica])
    .ListadoDeAmigos = NuloANada(rsTablaUsuarios![ListadoDeAmigos])
    .Password = NuloANada(rsTablaUsuarios![Password])
    .MensajesOffline = NuloANada(rsTablaUsuarios![MensajesOffline])
    .UsuarioBloqueado = rsTablaUsuarios![UsuarioBloqueado]
    .ApellidoYNombre = NuloANada(rsTablaUsuarios![ApellidoYNombre])
    .UltimoLogueo = NuloANada(rsTablaUsuarios![UltimoLogueo])
   End With
   rsTablaUsuarios.MoveNext
  Loop
  ' **************************************************************
  ' Emite un Warning si la Cantidad de Usuarios Registrados es
  ' Mayor a la cantidad de Usuarios Permitidos...
  ' **************************************************************
  If CInt(Configuracion.UsuariosSoportados) < CInt(Configuracion.CantidadDeUsuarios) Then
   EscribirEvento "The System Support " & Configuracion.UsuariosSoportados & ", But The System Have  " & _
                   CantidadDeRecords & " User's in the Database, For this Reason Some User's May Be can't Logued-In...", vbRed
  End If
  ' **************************************************************
  
 End If
 
 ' **************************************************************
 ' Informa la Cantidad de Usuarios Cargados en Memoria
 ' **************************************************************
 EscribirEvento "The System Load [" & CantidadDeRecords & "] User's in Memory...", vbBlue
 
End Function
Function GrabarLog(UsuarioNumero As Integer, Operacion As String) As Integer
' Devuelve:
'           1 : Todo OK
'           0 : No Se Grabo

 ' **************************************************************
 ' Nuevo Recorset...
 ' **************************************************************
 rsTablaLogs.AddNew

 ' **************************************************************
 ' Realiza el update del Password
 ' **************************************************************
 With Usuarios(UsuarioNumero)
  rsTablaLogs![IDNumericoUsuario] = UsuarioNumero
  rsTablaLogs![IDAliasUsuario] = .IDAliasUsuario
  rsTablaLogs![ApellidoYNombre] = .ApellidoYNombre
  rsTablaLogs![DireccionDeEmail] = .DireccionDeEmail
  rsTablaLogs![FechaDeNacimiento] = .FechaDeNacimiento
  rsTablaLogs![Edad] = .Edad
  rsTablaLogs![EstadoCivil] = .EstadoCivil
  rsTablaLogs![Humor] = .Humor
  rsTablaLogs![IDAliasUsuario] = .IDAliasUsuario
  rsTablaLogs![Intencion] = .Intencion
  rsTablaLogs![Ocupacion] = .Ocupacion
  rsTablaLogs![OtraInfo] = .OtraInfo
  rsTablaLogs![Sexo] = .Sexo
  rsTablaLogs![SigNo] = .SigNo
  rsTablaLogs![Telefono] = .Telefono
  rsTablaLogs![UbicacionGeografica] = .UbicacionGeografica
  rsTablaLogs![ListadoDeAmigos] = .ListadoDeAmigos
  rsTablaLogs![Password] = .Password
  rsTablaLogs![MensajesOffline] = .MensajesOffline
  rsTablaLogs![UsuarioBloqueado] = .UsuarioBloqueado
  rsTablaLogs![Password] = .Password
  rsTablaLogs![Operacion] = Operacion
  rsTablaLogs![OperacionFechaYHora] = Varios.FechaActual
 End With
 rsTablaUsuarios.Update
 
 GrabarLog = 1
 
End Function
Public Function ArreglarIndices()
Dim Contador, Numero As Integer
Dim Existe() As Boolean

 ' **************************************************************
 ' Definir para Verificar si existe...
 ' **************************************************************
 ReDim Existe(Configuracion.UsuariosSoportados)
 ' Pone en False todo...
 For Contador = 1 To Configuracion.UsuariosSoportados
  Existe(Contador) = False
 Next
 Do Until rsTablaUsuarios.EOF()
  Numero = rsTablaUsuarios![IDNumericoUsuario]
  Existe(Numero) = True
  rsTablaUsuarios.MoveNext
 Loop
 
 ' **************************************************************
 ' Segun lo que detecta, crea los indices faltantes...
 ' **************************************************************
 For Contador = 1 To Configuracion.UsuariosSoportados
  If Existe(Contador) = False Then
   rsTablaUsuarios.AddNew
   rsTablaUsuarios![IDNumericoUsuario] = Contador
   rsTablaUsuarios.Update
  End If
 Next
 
End Function
Public Function AgregarMensajeOffline(Para As String, Paquete As String, De As String) As Integer

 ' **************************************************************
 ' Agrega el Nuevo registro...
 ' **************************************************************
 With Variables.rsMensajeOffLine
  .AddNew
  ![UsuarioEmisor] = De
  ![FechaYHora] = Varios.FechaActualFormatoOffLine
  ![UsuarioPara] = Para
  ![Mensaje] = Paquete
  .Update
 End With
 
End Function
