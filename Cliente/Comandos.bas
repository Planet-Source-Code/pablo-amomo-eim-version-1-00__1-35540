Attribute VB_Name = "Comandos"
' **************************************************************
' En este modulo se procesan los comandos recibidos del Cliente
' **************************************************************
'   Comando 0:      Paquetes de Login
'   Comando 1:      Cambio de Estado del Usuario
Option Explicit
Function ComandoAccion_0(Datos As Variant) As String
 ' **************************************************************
 ' Formato del Paquete de Loguin...
 '   Comando Loguin 0: Validacion
 '                              1 - El Usuario No Existe
 '                              2 - El Password es Incorrecto
 '                              3 - Loguin Correcto...
 '                              4 - Usuario Lockeado
 '   Comando Loguin 1: Cambio de Password
 '                              1 - Ok
 ' **************************************************************
 
 ' **************************************************************
 ' Variables...
 ' **************************************************************
 Dim Comando, Resto, Sexo As String
 
 ' **************************************************************
 ' Separa el Comando de Loguin
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 Resto = Mid$(Datos, 2, 1)
  
 Select Case Comando
  Case 0:
    Configuracion.Logueado = 0
    If Resto = "1" Then Configuracion.Logueado = 1 ' Error : El Usuario No Existe
    If Resto = "2" Then Configuracion.Logueado = 2 ' Error : La Password es Incorrecta
    If Resto = "3" Then ' Ok    : Validacion Correcta
     Configuracion.Logueado = 3
     ' Define el Sexo del Usuario actual
     If Len(Datos) >= 53 Then
      Sexo = Mid$(Datos, 3, 1)
     End If
     If UCase(Sexo) <> "M" And UCase(Sexo) <> "F" Then
       Sexo = "M"
     End If
     Configuracion.Sexo = Sexo
     Configuracion.MiNombreYApellido = Mid$(Datos, 4)
    End If
    If Resto = "4" Then Configuracion.Logueado = 4  ' Usuario Lockeado
  Case 1:
    If Resto = "1" Then CambioDePasswordOk = True
  Case 2:
    ComandosProceso.ConfirmarElEnvioDeMail (Resto)
 End Select
 
End Function
Function ComandoAccion_1(Datos As Variant) As String
 ' **************************************************************
 ' Formato del Paquete de Estados...
 '   Comando Estado 0: Cambiar Estado
 '                              0 - error
 '                              1 - Ok
 '   Comando Estado 1: Cambio de Password
 ' **************************************************************
 
 ' **************************************************************
 ' Variables...
 ' **************************************************************
 Dim Comando, Resto As String

 ' **************************************************************
 ' Separa el Comando de Loguin
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 Resto = Mid$(Datos, 2)
 
 Select Case Comando
  Case 0:
   If Resto = 1 Then
    ' Si se confirma cambia el Estado del Usuario
    Configuracion.EstadoDelUsuario = NuevoEstadoUsuario.Numero
    Configuracion.EstadoActualTexto = NuevoEstadoUsuario.texto
    ' Graba el cambio de Estado
    Inicializar.UltimoEstado "Grabar", CStr(Configuracion.IDAliasUsuario)
    ' Lo toma como Logueado...
    CambiarEstadoDelCliente (3)
   End If
  Case 1: ' Procesa el Estado de un Amigo X...
   ComandosProceso.ProcesarEstadoDeUnAmigo (Resto)
 End Select
 
End Function
Function ComandoAccion_2(Datos As Variant) As String
 ' **************************************************************
 ' Formato del Paquete de Estados...
 '   Comando Estado 0: Datos de Un Usuario
 '                              0 - error
 '                              1 - Ok
 '                              2 - Usuario Incorrecto
 ' **************************************************************
 
 ' **************************************************************
 ' Variables
 ' **************************************************************
 Dim Comando, Resto As String
 
 ' **************************************************************
 ' Separa el Comando de Loguin
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 Resto = Mid$(Datos, 2)
 
 Select Case Comando
  Case 0:
  ' **************************************************************
   ' La Peticion de Datos tuvo un Error
   If Mid$(Resto, 1, 1) = "0" Then
    ' Aca no hace nada ya que el error se muestra por
    ' Time OUT en el formulario
   End If
   If Mid$(Resto, 1, 1) = "1" Or Mid$(Resto, 1, 1) = "2" Then
    ComandosProceso.ProcesarDatosUsuario Mid$(Resto, 1, 1), Mid$(Resto, 2)
   End If
  ' **************************************************************
  Case 1:
  ' **************************************************************
    ComandosProceso.ProcesarGrabacionUsuario Mid$(Resto, 1, 1), Mid$(Resto, 1, 1)
  ' **************************************************************
  Case 2:
  ' **************************************************************
   ComandosProceso.ProcesarPeticionDeListadoDeAmigos (Resto)
  Case 3:
  ' **************************************************************
   ComandosProceso.ProcesarGrabacionDeListadoDeAmigos (Resto)
  Case 4:
  ' **************************************************************
   ComandosProceso.ProcesarValidarAmigo (Resto)
  Case 5:
  ' **************************************************************
   ComandosProceso.ProcesarBusquedaAmigos (Resto)
 End Select
 
End Function
Function ComandoAccion_3(Datos As Variant) As String
Dim Respuesta As Integer

 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Datos) < 1 Then
  ' Descarta el Paquete...
  Exit Function
 End If
 
 ' Procesa el Intercambio de Paquetes, la loguica la deja en el Modulo siguiente...
 Respuesta = ComandosIntercambioDePaquete.IntercambioDePaquete(Datos)

End Function
Function ComandoAccion_4(Datos As Variant) As String
Dim Comando, Resto, Respuesta As String
 
 ' **************************************************************
 ' Valida el Paquete
 ' **************************************************************
 If Len(Datos) < 2 Then
  ' Descarta el Paquete...
  Exit Function
 End If
 
 ' **************************************************************
 ' Define la Accion
 ' **************************************************************
 Comando = Mid$(Datos, 1, 1)
 Resto = Mid$(Datos, 2)
 Select Case Comando
  Case "M" ' Intercambio de Mensaje
   Respuesta = ComandosIntercambioDePaquete.IntercambioDeMensaje(Resto)
 End Select
 
End Function
