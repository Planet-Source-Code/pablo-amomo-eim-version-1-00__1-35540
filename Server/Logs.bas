Attribute VB_Name = "Logs"
Option Explicit
Sub EscribirEvento(Evento As String, Color As Variant)

  ' Va al Final...
  Server.MensajesServidor.SelStart = Len(Server.MensajesServidor.Text) + 1
  ' Setea el Color Negro como Original
  Server.MensajesServidor.SelColor = vbBlack
  ' Escribe la Hora del Evento
  Server.MensajesServidor.SelText = Now & ":"
  ' Setea el Color Definido por el Usuario
  Server.MensajesServidor.SelColor = Color
  ' Escribe el Evento
  Server.MensajesServidor.SelText = " " & Evento & vbCrLf
  ' Devuelve el Color Original
  Server.MensajesServidor.SelColor = vbBlack
  ' Va al Final...
  Server.MensajesServidor.SelStart = Len(Server.MensajesServidor.Text) + 1

End Sub
