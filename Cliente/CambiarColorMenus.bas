Attribute VB_Name = "CambiarColorMenus"
Option Explicit
Sub CambiarColorTreeList(Color As Long)
 
 ' **************************************************************
 ' Cambia el Color del Listado de Amigos
 ' **************************************************************
 Call SendMessage(Cliente.ListadoDeAmigos.hwnd, 4381&, 0, ByVal Color)
 
End Sub
