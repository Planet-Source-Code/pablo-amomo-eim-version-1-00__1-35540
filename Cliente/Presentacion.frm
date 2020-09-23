VERSION 5.00
Begin VB.Form Presentacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   ClipControls    =   0   'False
   Icon            =   "Presentacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Presentacion.frx":08CA
   ScaleHeight     =   2085
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4050
      Top             =   2430
   End
   Begin VB.Image Touch 
      Height          =   2085
      Left            =   0
      MouseIcon       =   "Presentacion.frx":2FD8
      MousePointer    =   99  'Custom
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "Presentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **************************************************************
' Variables Compartidas de Formulario (Para Busqueda)
' **************************************************************
Public FormularioNombre As String
Public AliasUsuario, LabelOk As String
Public Refresco As Boolean
Public GraboLosDatosUsuario As Boolean
Public CambioDeDatosUsuario As Boolean
Public Grabando, Refrescando As Boolean
' **************************************************************
' **************************************************************
Private Sub Form_Load()

 ' **************************************************************
 ' Espera un Segundo...
 ' **************************************************************
 Me.FormularioNombre = "Presentacion"
 TimeOut.Enabled = True
  
End Sub
Private Sub TimeOut_Timer()

 ' **************************************************************
 ' Despues de 1 Segundo Sale del Form...
 ' **************************************************************
 If TimeOut Then
  Unload Me
 End If
 
End Sub
Private Sub Touch_Click()

 Unload Me
 
End Sub
