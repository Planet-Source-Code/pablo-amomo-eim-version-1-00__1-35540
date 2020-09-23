Attribute VB_Name = "Declaraciones"
Option Explicit

' **************************************************************
' Declaracion Para la Ejecucion de Sonidos...
' **************************************************************
Public Declare Function waveOutGetNumDevs Lib "WINMM.DLL" () As Long
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

' **************************************************************
' Declaracion Para la llamada al IExplorer desde Mensajes, Mail,
' etc.
' **************************************************************
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' **************************************************************
' Usado para el Drag And Drop de las Ventanas
' **************************************************************
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
' **************************************************************

' //////

' **************************************************************
' Constantes usadas por el IconTray
' **************************************************************
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
' **************************************************************
