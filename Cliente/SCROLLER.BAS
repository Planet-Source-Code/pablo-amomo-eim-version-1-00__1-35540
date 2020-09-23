Attribute VB_Name = "Scroller"
#If Win32 Then
    Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
#Else
    Private Declare Function PutFocus% Lib "user" Alias "SetFocus" (ByVal hWd%)
    Private Declare Function SendMessage& Lib "user" (ByVal hWd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
#End If

Function ScrollText&(TextBox As Control, vLines As Integer)
    ' The Windows Version Stuff...
    #If Win32 Then
        Dim Success As Long
        Dim SavedWnd As Long
        Dim R As Long
    #Else
        Dim Success As Integer
        Dim SavedWnd As Integer
        Dim R As Integer
    #End If
    
    Const EM_LINESCROLL = &HB6
    
    SavedWnd = Screen.ActiveControl.hwnd
    Lines& = vLines
    
    ' Scroll the lines, using the SendMessage.
    Success = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)
    
    ' Restore the focus to the original control which had
    ' the previous focus, which we previously recorded
    ' before.
    R = PutFocus(SavedWnd)
    
    ' Return the number of lines actually scrolled.
    ScrollText& = Success
End Function
Function ScrollTreeView&(ArbolView As TreeView, Direccion As String)

  Select Case Direccion
   Case "Arriba"
     SendMessage ArbolView.hwnd, 277&, 0&, vbNull
   Case "Abajo"
     SendMessage ArbolView.hwnd, 277&, 1&, vbNull
   Case "Derecha"
     SendMessage ArbolView.hwnd, 276&, 1&, vbNull
   Case "Izquierda"
     SendMessage ArbolView.hwnd, 276&, 0&, vbNull
  End Select

End Function
Function ScrollListBox&(ListadoBox As ListBox, Direccion As String)

  Select Case Direccion
   Case "Arriba"
     SendMessage ListadoBox.hwnd, 277&, 0&, vbNull
   Case "Abajo"
     SendMessage ListadoBox.hwnd, 277&, 1&, vbNull
   Case "Derecha"
     SendMessage ListadoBox.hwnd, 276&, 1&, vbNull
   Case "Izquierda"
     SendMessage ListadoBox.hwnd, 276&, 0&, vbNull
  End Select

End Function

