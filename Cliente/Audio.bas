Attribute VB_Name = "Audio"
' **************************************************************
' Sonido 001: Sonido de Cambio De Estado Abrir
' Sonido 002: Sonido de Cambio De Estado Cerrar
' Sonido 003: Click
' Sonido 004: Aviso de llegada de Mensajes Pendientes
' **************************************************************
Option Explicit
Public Function TienePlacaDeSonido() As Boolean
 
 ' **************************************************************
 ' Valida si tiene placa de sonido
 ' **************************************************************
 If waveOutGetNumDevs() > 0 Then
   TienePlacaDeSonido = True
  Else
   TienePlacaDeSonido = False
 End If
 
End Function
Public Function EjecutarSonido(Sonido As String, Optional SioSi As String)

 ' **************************************************************
 ' Verifica si el Usuario desea recibir Sonidos
 ' **************************************************************
 If Not Configuracion.SonidoActivado Then
  If UCase(SioSi) <> "SI" Then Exit Function
 End If
 
 ' **************************************************************
 ' Ejecuta el Sonido, sino tiene placa emite un BEEP
 ' **************************************************************
 If TienePlacaDeSonido Then
   Select Case CInt(Sonido)
    Case 1:
     Call sndPlaySound(BufferSonido1(0), 1 Or 2 Or &H4)
    Case 2:
     Call sndPlaySound(BufferSonido2(0), 1 Or 2 Or &H4)
    Case 3:
     Call sndPlaySound(BufferSonido3(0), 1 Or 2 Or &H4)
    Case 4:
     Call sndPlaySound(BufferSonido4(0), 1 Or 2 Or &H4)
    Case 5:
     Call sndPlaySound(BufferSonido5(0), 1 Or 2 Or &H4)
   End Select
  Else
   Beep
 End If
   
End Function
