VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const NIM_ADD = &H0 'Add a new icon to the Systray
Private Const NIM_MODIFY = &H1 'Modify an existing icon
Private Const NIM_DELETE = &H2 'Delete an existing icon
Private Const NIF_MESSAGE = &H1 'Tell that the message has been updated
Private Const NIF_ICON = &H2 'Tell that the icon picture has been changed
Private Const NIF_TIP = &H4 'Tell that a new ToopTip for the icon is set
Private Const WM_MOUSEMOVE = &H200 'Used as the ID of the callback message
Private Const MAX_TIP_LENGTH As Long = 64 'This is the max length
Private Type NOTIFYICONDATA
    cbSize As Long 'The size of this type
    hwnd As Long 'The hWnd that will receive the CallBack message
    uId As Long 'The ID of the application. Zero represent this application
    uFlags As Long 'The flags. Look at the constants beginning by NIF_ for the flags and their definition
    uCallbackMessage As Long 'This is the callback message
    hIcon As Long 'An handle to the icon that will be displayed
    szTip As String * MAX_TIP_LENGTH 'The string of the ToopTip. Must be terminating by a null zero (chr(0))
End Type
Private nidTrayIcon As NOTIFYICONDATA
Private bIconDisplayed As Boolean
Private bUpdateOnChange As Boolean
Public Event NIError(ByVal ErrorNumber As Long)
Public Function Initialize(ByVal hwnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
  
  With nidTrayIcon
   .cbSize = Len(nidTrayIcon)
   .hIcon = hIcon
   .hwnd = hwnd
   .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
   .uCallbackMessage = uCallbackMessage
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uId = vbNull
  End With
  bIconDisplayed = False
  bUpdateOnChange = True

End Function
Public Sub CambiarIconoSystemTray(Icono As Long, Etiqueta As String)
 
 With nidTrayIcon
  .hIcon = Icono
  .szTip = Left(Etiqueta, MAX_TIP_LENGTH - 1) & vbNullChar
 End With
 ShowIcon
 
End Sub
Public Function ShowIcon() As Boolean

  If Not bIconDisplayed Then 'If the icon is not already displayed...
    ShowIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
    If ShowIcon = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    Else
      bIconDisplayed = True 'The icon is displayed
    End If
   Else
    ShowIcon = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
  End If
  
End Function
Public Function HideIcon() As Boolean

  If bIconDisplayed Then 'If the icon is displayed...
    HideIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
    If HideIcon = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    Else
      bIconDisplayed = False 'The icon is not displayed
    End If
  End If
End Function

Public Property Let IconHandle(ByVal hIcon As Long)
  'Change the icon displayed in the systray for the icon handled
  'by hIcon. bIconStatus must be true.
  nidTrayIcon.hIcon = hIcon
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_ICON
    Update 'Make the icon change appear
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let TipText(ByVal sTip As String)
  'Change the tooltip text of the icon by sTip
  'Remove the last character of it length is MAX_TIP_LENGTH because
  'we need to include a Null Zero
  nidTrayIcon.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_TIP
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
  'Change the callback message that is send to the form.
  'The new message is uCallbackMessage.
  nidTrayIcon.uCallbackMessage = uCallbackMessage
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_MESSAGE
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Function Update() As Boolean
  'Update the icon on the systray. Call this function when you
  'want to refresh the icon or when you have made a change to
  'a property when bUpdateOnChange is False.
  'Also called internaly when bUpdateOnChange is true.
  If bIconDisplayed Then 'The icon must be showed to make change
    Update = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
    If Update = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    End If
  End If
End Function

Public Property Get IconHandle() As Long
  'Return the handle of the current icon
  IconHandle = nidTrayIcon.hIcon
End Property

Public Property Get TipText() As String
  'Return the tooltip text w/o the ending null zero
  TipText = Left(nidTrayIcon.szTip, Len(nidTrayIcon.szTip) - 1)
End Property

Public Property Get CallbackMessage() As Long
  'Return the current callback message
  CallbackMessage = nidTrayIcon.uCallbackMessage
End Property

Public Property Let UpdateOnChange(bUpdate As Boolean)
  'Set the variable bUpdateOnChange to a new value. If bUpdate is
  'True, then when a change to a properties of nidTrayIcon is made,
  'Shell_NotifyIcon is call. Otherwise, the function Update must
  'be called to make the modification appear.
  bUpdateOnChange = bUpdate
End Property

Private Property Get UpdateOnChange() As Boolean
  'Return True if bUpdateOnChange is True
  UpdateOnChange = bUpdateOnChange
End Property

Private Sub Class_Terminate()
  'The destructor of the class. It remove the icon from the Systray
  HideIcon
End Sub

Public Property Get Visible() As Boolean
  'Return if the icon in the systray is visible
  If bIconDisplayed Then
    Visible = True
  End If
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
  'This property does the same thing than calling ShowIcon() and
  'HideIcon()
  If bVisible Then
    ShowIcon
  Else
    HideIcon
  End If
End Property
