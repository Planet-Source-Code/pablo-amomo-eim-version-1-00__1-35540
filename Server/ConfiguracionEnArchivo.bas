Attribute VB_Name = "ConfiguracionEnArchivo"
Option Explicit
Sub CargarConfiguracionArchivo()

 ' **********************************************************************
 ' Graba el Archivo
 ' **********************************************************************
 Open App.Path & "\EIM.cfg" For Random As #1 Len = Len(ArchivoConfiguracion)
  Get #1, 1, ArchivoConfiguracion
 Close #1
 
 ' **********************************************************************
 ' Carga Configuracion levantada del Archivo...
 ' **********************************************************************
 With Configuracion
  .UsuariosSoportados = ArchivoConfiguracion.UsuariosSoportados
  .UbicacionBaseDeDatos = ArchivoConfiguracion.UbicacionBaseDeDatos
  .NombreDeLaBaseDeDatos = ArchivoConfiguracion.NombreDeLaBaseDeDatos
  .PortTCP = ArchivoConfiguracion.PortTCP
  .DireccionIPSMTP = ArchivoConfiguracion.DireccionIPSMTP
  .DireccionEMAILAdministrador = ArchivoConfiguracion.DireccionEMAILAdministrador
  .PermitirCrear = ArchivoConfiguracion.PermitirCrear
 End With
 
 
 ' **********************************************************************
 ' Verifica que los Valores en el Archivo sean Correctos, si no los
 ' son los arregla y graba el Archivo de Configuracion Nuevamente...
 ' **********************************************************************
 With Configuracion
  If IsNumeric(.UsuariosSoportados) = False Then
   .UsuariosSoportados = 300
   GrabarConfiguracionArchivo
  End If
  If IsNull(.UbicacionBaseDeDatos) Or Trim(.UbicacionBaseDeDatos) = "" Or Asc(Mid$(.UbicacionBaseDeDatos, 1, 1)) = 0 Then
   .UbicacionBaseDeDatos = App.Path & "\DataBase\"
   GrabarConfiguracionArchivo
  End If
  If IsNull(.NombreDeLaBaseDeDatos) Or Trim(.NombreDeLaBaseDeDatos) = "" Or Asc(Mid$(.NombreDeLaBaseDeDatos, 1, 1)) = 0 Then
   .NombreDeLaBaseDeDatos = "EIM.MDB"
   GrabarConfiguracionArchivo
  End If
  If IsNumeric(.PortTCP) = False Then
   .PortTCP = 24157
   GrabarConfiguracionArchivo
  End If
  If .PortTCP = 0 Then
   .PortTCP = 24157
   GrabarConfiguracionArchivo
  End If
  If IsNull(.DireccionIPSMTP) Or Trim(.DireccionIPSMTP) = "" Or Asc(Mid$(.DireccionIPSMTP, 1, 1)) = 0 Then
   .DireccionIPSMTP = "127.0.0.1"
   GrabarConfiguracionArchivo
  End If
  If IsNull(.DireccionEMAILAdministrador) Or Trim(.DireccionEMAILAdministrador) = "" Or Asc(Mid$(.DireccionEMAILAdministrador, 1, 1)) = 0 Then
   .DireccionEMAILAdministrador = "Mail@Mail.com"
   GrabarConfiguracionArchivo
  End If
  If .PermitirCrear <> True And .PermitirCrear <> False Then
   .PermitirCrear = False
  End If
 End With
 ' **********************************************************************
 ' **********************************************************************
 ' Se se cargaron mas de 300 usuarios concurrentes, graba 300
 ' **********************************************************************
 If Configuracion.UsuariosSoportados > 300 Then
  Configuracion.UsuariosSoportados = 300
  ' Grabar Configuracion
  GrabarConfiguracionArchivo
 End If
 
 
End Sub
Sub GrabarConfiguracionArchivo()

 ' **********************************************************************
 ' Carga las Variables a Grabar en el Archivo de Configuracion
 ' **********************************************************************
 With ArchivoConfiguracion
  .DireccionIPSMTP = Trim(Configuracion.DireccionIPSMTP)
  .NombreDeLaBaseDeDatos = Trim(Configuracion.NombreDeLaBaseDeDatos)
  .PermitirCrear = Configuracion.PermitirCrear
  .PortTCP = Configuracion.PortTCP
  .UbicacionBaseDeDatos = Trim(Configuracion.UbicacionBaseDeDatos)
  .DireccionEMAILAdministrador = Trim(Configuracion.DireccionEMAILAdministrador)
  .UsuariosSoportados = Trim(Configuracion.UsuariosSoportados)
 End With
 
 ' **********************************************************************
 ' Graba el Archivo
 ' **********************************************************************
 Open App.Path & "\EIM.cfg" For Random As #1 Len = Len(ArchivoConfiguracion)
  Put #1, 1, ArchivoConfiguracion
 Close #1

End Sub
