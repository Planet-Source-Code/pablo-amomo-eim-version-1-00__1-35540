VERSION 5.00
Begin VB.Form IcoMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "IcoMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Image MenuDesplegable 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   480
      Picture         =   "IcoMenu.frx":000C
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line LineaBlanca 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   0
      X2              =   840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Recuadro 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   300
      X2              =   300
      Y1              =   0
      Y2              =   3800
   End
   Begin VB.Line LineaSeparacion 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   400
      X2              =   3000
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Image IMenu 
      Height          =   240
      Index           =   0
      Left            =   30
      Picture         =   "IcoMenu.frx":0387
      Stretch         =   -1  'True
      Top             =   30
      Width           =   240
   End
   Begin VB.Label LBLmenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   0
      Top             =   1170
      Width           =   45
   End
   Begin VB.Shape ItemFocus 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00FF8080&
      BorderStyle     =   0  'Transparent
      Height          =   285
      Index           =   0
      Left            =   300
      Top             =   0
      Width           =   2235
   End
   Begin VB.Shape SombraIzquierda 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2355
      Left            =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Image GrupoGrisado 
      Height          =   825
      Left            =   1920
      Picture         =   "IcoMenu.frx":04D1
      Top             =   1350
      Width           =   1110
   End
End
Attribute VB_Name = "IcoMenu"
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

Public DesdeMenu As String
' **************************************************************

' **************************************************************
' Para Desplegable Colgado
' **************************************************************
Private DesplegableColgado() As Boolean
Private HandleVentanaOrigen As Long
Public SeLlamoAlDesplegable As Boolean
Public CantidadDeItems As Integer

Option Explicit

Private LastIco As Long     ' Used for Switch Off Item
Public MaxWidth As Long    ' Used for Calcule the Form Width
Const MenuHeight = 270      ' Alto Del Menu (Cada Item)
Const MargenSuperior = 15   ' Define el Margen Superior
    
' **************************************************************
' Eventos Raiseados...
' **************************************************************
Public Event Click(ByVal Index As Long, Tag As String)
Public Event Closed()
Public Event MouseDown(ByVal Index As Long, Tag As String, Button As Integer, Shift As Integer)
Public Event MouseUp(ByVal Index As Long, Tag As String, Button As Integer, Shift As Integer)
Public Event MouseMove(ByVal Index As Long, Tag As String)
Public Sub SetItem(ByVal Index As Long, ByVal Caption As String, Optional Icon As IPictureDisp, Optional key As String, Optional DesplegableyColgado)
On Error Resume Next
    
    Load LBLmenu(Index)
    Load IMenu(Index)
    Load LineaSeparacion(Index)
    Load ItemFocus(Index)
    
    ' ******************************************************************
    ' Carga los Desplegables Colgados
    ' ******************************************************************
    ReDim Preserve DesplegableColgado(Index)
    If DesplegableyColgado = False Or DesplegableyColgado = vbNull Then
      DesplegableColgado(Index) = False
     Else
      DesplegableColgado(Index) = True
    End If
    
    ' ******************************************************************
    ' Pone la Imagen del Menu
    ' ******************************************************************
    If Not Icon Is Nothing Then
     Set IMenu(Index).Picture = Icon
    End If
    
    ' ******************************************************************
    ' Pone los Labels
    ' ******************************************************************
    LBLmenu(Index).Caption = Caption & Space(10)
    LBLmenu(Index).Left = 450
    LBLmenu(Index).ForeColor = Variables.FontMenuDescolgableAbiertoFont
    If LBLmenu(Index).WidtH + 300 > MaxWidth Then MaxWidth = LBLmenu(Index).WidtH + 300
    
    ' ******************************************************************
    ' Si en el Label no Hay nada pone una Linea de Separacion
    ' ******************************************************************
    If Trim(LBLmenu(Index).Caption) = "" Then
     With LineaSeparacion(Index)
      .X1 = 450
      .X2 = IcoMenu.WidtH - 450
      .Y1 = (MenuHeight * Index) + MargenSuperior + 130
      .Y2 = (MenuHeight * Index) + MargenSuperior + 130
      .BorderColor = Variables.FontMenuDescolgableAbiertoFranjas
      .Visible = True
     End With
    End If
    
    ' ******************************************************************
    ' Pone el Label Visible o no, segun si tiene algo en Caption...
    ' ******************************************************************
    With LBLmenu(Index)
      .Tag = key
      .Top = (MenuHeight * Index) + MargenSuperior + 40
     If Trim(LBLmenu(Index).Caption) <> "" Then
       .Visible = True
      Else
       .Visible = False
     End If
    End With
    
    ' ******************************************************************
    ' Muestra el Icono (Grafico) en el Menu
    ' ******************************************************************
    With IMenu(Index)
      .Top = (MenuHeight * Index) + 30 + MargenSuperior
      .Left = (390 - .WidtH) / 2
      If Trim(LBLmenu(Index).Caption) <> "" Then
        .Visible = True
        .ZOrder (0)
       Else
        .Visible = False
        .ZOrder (0)
      End If
    
    ' ******************************************************************
    ' Pone el Item que simula el Foco...
    ' ******************************************************************
    With ItemFocus(Index)
      .Top = (MenuHeight * Index) + MargenSuperior + 15
      .Left = 300
      .WidtH = 3100
      If Trim(LBLmenu(Index).Caption) <> "" Then
        .Visible = True
       Else
        .Visible = False
      End If
    End With
   
      ' ******************************************************************
      ' Pone la Imagen de Desplegable...
      ' ******************************************************************
      If DesplegableColgado(Index) = True Then
       Load MenuDesplegable(Index)
       With MenuDesplegable(Index)
         .Top = (MenuHeight * Index) + 30 + MargenSuperior
         .Left = Me.WidtH - .WidtH - 1245  ' ((390 - .Width) / 2)
         .Visible = True
         .ZOrder (0)
       End With
      End If
    End With
       
    ' ******************************************************************
    ' Define la Cantidad Total de Items...
    ' ******************************************************************
    If Me.CantidadDeItems < Index Then
     CantidadDeItems = Index
    End If
        
End Sub
Public Sub DestroyMenu()
    
    ' ******************************************************************
    ' Descarga el Formulario...
    ' ******************************************************************
    Unload Me

End Sub

Public Sub ShowMenu(Optional ByVal Left As Long = -1, Optional ByVal Top As Long = -1, Optional Largo As Long, Optional HandleAnterior As Long, Optional LadoIzqoDer)
    Dim i As Long
    Dim CurPos As POINTAPI
    
    ' **************************************************************
    ' Verifica que no haya ningun MSGBOX o INPUT BOX Abierto...
    ' Si lo hay sale sin hacer nada...
    ' **************************************************************
    Dim Contador As Integer
    For Contador = 0 To Forms.Count - 1
     DoEvents
     If Forms(Contador).FormularioNombre = "MensajesBox" Then
      If Forms(Contador).Modal Then Exit Sub
     End If
     If Forms(Contador).FormularioNombre = "IngresoBox" Then
      If Forms(Contador).Modal Then Exit Sub
     End If
    Next
    
    ' ******************************************************************
    ' Define si viene de otro menu
    ' ******************************************************************
    If HandleAnterior <> 0 Then
      HandleVentanaOrigen = HandleAnterior
     Else
      HandleVentanaOrigen = -1
    End If
    
    ' ******************************************************************
    ' Define el Alto Total del Menu (Formulario)
    ' ******************************************************************
    Me.HeighT = MenuHeight * LBLmenu.Count + 70
        
    ' ******************************************************************
    ' Define la Posicion donde Se Mostrara el Menu.
    ' ******************************************************************
    If Left = -1 Then
      GetCursorPos CurPos
      Me.Left = CurPos.X * Screen.TwipsPerPixelX
      Me.Top = CurPos.Y * Screen.TwipsPerPixelY
    Else
      Me.Left = Left
      Me.Top = Top + 30
    End If
    For i = LBLmenu.LBound To LBLmenu.UBound
      With LBLmenu(i)
         .AutoSize = False
         .WidtH = MaxWidth
         .HeighT = MenuHeight
      End With
    Next
    Me.WidtH = MaxWidth + 90
    
    ' ******************************************************************
    ' Arregla la Posicion del icono Desplegable...
    ' ******************************************************************
    For i = 0 To Me.CantidadDeItems
     If DesplegableColgado(i) = True Then
      With MenuDesplegable(i)
       .Left = Me.WidtH - .WidtH
      End With
     End If
    Next
    
    ' ******************************************************************
    ' Define el Recuadro del Formulario
    ' ******************************************************************
    Me.Recuadro.Top = 0
    Me.Recuadro.Left = 0
    Me.Recuadro.WidtH = Me.WidtH
    Me.Recuadro.HeighT = Me.HeighT
        
    ' ******************************************************************
    ' Define la SombraIzquierda
    ' ******************************************************************
    Me.SombraIzquierda.Top = 0
    Me.SombraIzquierda.Left = 0
    Me.SombraIzquierda.WidtH = 390
    Me.SombraIzquierda.HeighT = Me.HeighT
        
    ' ******************************************************************
    ' Muestra la Linea Blanca
    ' ******************************************************************
    Me.LineaBlanca.X1 = 10
    Me.LineaBlanca.X2 = Largo - 20
    Me.LineaBlanca.Y1 = 1
    Me.LineaBlanca.Y2 = 1
    If IsNull(Largo) Or Largo = 0 Then
      Me.LineaBlanca.Visible = False
     Else
      Me.LineaBlanca.Visible = True
    End If
    
    ' ******************************************************************
    ' Define la Linea Oscura...
    ' ******************************************************************
    Me.Line1.X1 = Me.SombraIzquierda.WidtH - 10
    Me.Line1.X2 = Me.SombraIzquierda.WidtH - 10
    Me.Line1.Y1 = 10
    Me.Line1.Y2 = Me.HeighT
    Me.Line1.Visible = True
    Me.Line1.ZOrder 0
    
    ' ******************************************************************
    ' Manda al Fondo el Grupo Grisado...
    ' ******************************************************************
    Me.GrupoGrisado.ZOrder 1
    Me.GrupoGrisado.Left = Me.WidtH - Me.GrupoGrisado.WidtH + 40
    Me.GrupoGrisado.Top = Me.HeighT - Me.GrupoGrisado.HeighT
    
    ' ******************************************************************
    ' Color de la Franja de la Izquierda...
    ' ******************************************************************
    Me.SombraIzquierda.BackColor = Variables.FontMenuDescolgableAbiertoFranjas
    Me.SombraIzquierda.FillColor = Variables.FontMenuDescolgableAbiertoFranjas
    Me.LineaBlanca.BorderColor = Variables.FontMenuDescolgableAbiertoFranjas
    Me.Line1.BorderColor = Variables.FontMenuDescolgableAbiertoLineaOscura
    Me.BackColor = Variables.FontMenuDescolgableAbiertoFondoFormulario
    If Variables.FontMenuDescolgableAbiertoMostrarGrafico Then
      Me.GrupoGrisado.Visible = True
     Else
      Me.GrupoGrisado.Visible = False
    End If
    
    ' ******************************************************************
    ' Muestra el Menu
    ' ******************************************************************
    Me.Show
    
End Sub
Public Sub HideMenu()
On Error GoTo HideMenuError

    ' ******************************************************************
    ' Si habia un Menu descolgable y perdio el Foco, elimina el
    ' formulario
    ' ******************************************************************
    If DesplegableColgado(LastIco) Then
      RaiseEvent Click(LastIco, "PerdioFoco")
    End If
    
    ' Conectar
    Cliente.MenuBotonConectar.BackColor = Variables.FontFOndoMenuDescolgable
    Cliente.MenuBotonConectar.ForeColor = Variables.FontMenuDescolgable
    Cliente.MenuBotonConectar.FontUnderline = False
    Cliente.MenuBotonConectar.BorderStyle = 0
    Cliente.MenuBotonConectar.BackStyle = 0
    ' Case "Configuracion"
    Cliente.MenuBotonConfiguracion.BackColor = Variables.FontFOndoMenuDescolgable
    Cliente.MenuBotonConfiguracion.ForeColor = Variables.FontMenuDescolgable
    Cliente.MenuBotonConfiguracion.FontUnderline = False
    Cliente.MenuBotonConfiguracion.BorderStyle = 0
    Cliente.MenuBotonConfiguracion.BackStyle = 0
    ' Case "Amigos"
    Cliente.MenuBotonAmigos.BackColor = Variables.FontFOndoMenuDescolgable
    Cliente.MenuBotonAmigos.ForeColor = Variables.FontMenuDescolgable
    Cliente.MenuBotonAmigos.FontUnderline = False
    Cliente.MenuBotonAmigos.BorderStyle = 0
    Cliente.MenuBotonAmigos.BackStyle = 0
    ' Case "Ayuda"
    Cliente.MenuBotonAyuda.BackColor = Variables.FontFOndoMenuDescolgable
    Cliente.MenuBotonAyuda.ForeColor = Variables.FontMenuDescolgable
    Cliente.MenuBotonAyuda.FontUnderline = False
    Cliente.MenuBotonAyuda.BorderStyle = 0
    Cliente.MenuBotonAyuda.BackStyle = 0

    ' ******************************************************************
    ' Lo Esconde....
    ' ******************************************************************
    Me.Hide
     
    Exit Sub ' Sale !
    
HideMenuError:
    Resume Next
End Sub
Private Sub Form_Initialize()
    
    ' ******************************************************************
    ' Bandera de Desplegable...
    ' ******************************************************************
    SeLlamoAlDesplegable = False
    
    ' ******************************************************************
    ' Variables...
    ' ******************************************************************
    LastIco = 0
    MaxWidth = 90
    
    ' ******************************************************************
    ' Tipo de Formulario....
    ' ******************************************************************
    FormularioNombre = "VentanaMenu"
    
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_LostFocus()

    On Error GoTo ErrorIcoMenuFom
  
    If Me.Visible = True Then
     If HandleVentanaOrigen <> -1 Then
      Varios.PonerFocoEnVentana (HandleVentanaOrigen)
     End If
     ' Solo lo esconde si no es un desplegable...
     If DesplegableColgado(LastIco) <> True Then
      ' Solo lo descarga si no esta abierto el desplegable...
     End If
     If SeLlamoAlDesplegable = False Then
      Me.Hide
     End If
    End If
    Select Case Me.DesdeMenu
     Case "Coneccion"
      Cliente.MenuBotonConectar.BackColor = Variables.FontFOndoMenuDescolgable
      Cliente.MenuBotonConectar.ForeColor = Variables.FontMenuDescolgable
      Cliente.MenuBotonConectar.FontUnderline = False
      Cliente.MenuBotonConectar.BorderStyle = 0
      Cliente.MenuBotonConectar.BackStyle = 0
     Case "Configuracion"
      Cliente.MenuBotonConfiguracion.BackColor = Variables.FontFOndoMenuDescolgable
      Cliente.MenuBotonConfiguracion.ForeColor = Variables.FontMenuDescolgable
      Cliente.MenuBotonConfiguracion.FontUnderline = False
      Cliente.MenuBotonConfiguracion.BorderStyle = 0
      Cliente.MenuBotonConfiguracion.BackStyle = 0
     Case "Amigos"
      Cliente.MenuBotonAmigos.BackColor = Variables.FontFOndoMenuDescolgable
      Cliente.MenuBotonAmigos.ForeColor = Variables.FontMenuDescolgable
      Cliente.MenuBotonAmigos.FontUnderline = False
      Cliente.MenuBotonAmigos.BorderStyle = 0
      Cliente.MenuBotonAmigos.BackStyle = 0
     Case "Ayuda"
      Cliente.MenuBotonAyuda.BackColor = Variables.FontFOndoMenuDescolgable
      Cliente.MenuBotonAyuda.ForeColor = Variables.FontMenuDescolgable
      Cliente.MenuBotonAyuda.FontUnderline = False
      Cliente.MenuBotonAyuda.BorderStyle = 0
      Cliente.MenuBotonAyuda.BackStyle = 0
    End Select
         
    RaiseEvent Closed
    
ErrorIcoMenuFom:

End Sub

'===============================================================================
'
'===============================================================================
Private Sub IMenu_Click(Index As Integer)
    LBLmenu_Click Index
End Sub

'===============================================================================
'
'===============================================================================
Private Sub IMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LBLmenu_MouseMove(Index, Button, Shift, X, Y)
End Sub

'===============================================================================
'
'===============================================================================
Private Sub LBLmenu_Click(Index As Integer)
    
    
    If DesplegableColgado(Index) <> True Then
      Audio.EjecutarSonido "003"
      Me.Hide
      If HandleVentanaOrigen <> -1 Then
       Varios.DescargarVentanaHandle (HandleVentanaOrigen)
      End If
     Else
      SeLlamoAlDesplegable = True
    End If
    RaiseEvent Click(Index, LBLmenu(Index).Tag)
    
End Sub

'===============================================================================
'
'===============================================================================
Private Sub LBLmenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Index, LBLmenu(Index).Tag, Button, Shift)
End Sub

'===============================================================================
'
'===============================================================================
Private Sub LBLmenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' ******************************************************************
    ' Si el Item se encuentra Inactivo Sale sin hacer nada...
    ' ******************************************************************
    If LBLmenu(Index).Enabled = False Then Exit Sub
    
    ' ******************************************************************
    ' Si no hay ningun item seleccionado... Sale...
    ' ******************************************************************
    If LastIco = Index And Index <> 0 Then Exit Sub
    
    ' ******************************************************************
    ' Pone los Colores de las Letras de los Labels...
    ' ******************************************************************
    ' Vuelve a la Normalidad el Anterior
    LBLmenu(LastIco).ForeColor = Variables.FontMenuDescolgableAbiertoFont
    ' Pone el Nuevo Color
    LBLmenu(Index).ForeColor = Variables.FontMenuDescolgableAbiertoHighLightLetra
    ' ******************************************************************
    
    ' ******************************************************************
    ' Define el Marco remarcado del Label
    ' ******************************************************************
    ' Borra el Antiguo Focus
    ItemFocus(LastIco).BackStyle = 0
    ItemFocus(LastIco).BorderStyle = 0
    ' Crea el Nuevo Focus
    ItemFocus(Index).BackStyle = 1
    ItemFocus(Index).BackColor = Variables.FontMenuDescolgableAbiertoHighLightFondo
    ItemFocus(Index).Left = 400
    ItemFocus(Index).WidtH = Me.WidtH - 440
    ItemFocus(Index).BorderStyle = 1
    ItemFocus(Index).BorderColor = Variables.FontMenuDescolgableAbiertoHighLightBorde
    ' ******************************************************************
    
    ' ******************************************************************
    ' Si habia un Menu descolgable y perdio el Foco, elimina el
    ' formulario
    ' ******************************************************************
    If DesplegableColgado(LastIco) Then
      RaiseEvent Click(LastIco, "PerdioFoco")
    End If

    ' ******************************************************************
    ' Deja definido el Ultimo Item del Menu con Foco
    ' ******************************************************************
    LastIco = Index
    ' ******************************************************************
    
    ' ******************************************************************
    ' Si es desplegable y colgado lo Muestra al Toque...
    ' ******************************************************************
    If DesplegableColgado(Index) Then
     LBLmenu_Click (Index)
     SeLlamoAlDesplegable = True
     Exit Sub
    End If
    
    ' ******************************************************************
    ' Captura el Evento de Movimiento de Mouse sobre el Item
    ' ******************************************************************
    RaiseEvent MouseMove(Index, LBLmenu(Index).Tag)
    ' ******************************************************************

End Sub
Private Sub LBLmenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' ******************************************************************
    ' Envia un Evento....
    ' ******************************************************************
    RaiseEvent MouseUp(Index, LBLmenu(Index).Tag, Button, Shift)

End Sub
Public Sub HabilitarItem(Item As Integer, Estado As Boolean)

    ' ******************************************************************
    ' Habilitar / Deshabilita un Item...
    ' ******************************************************************
    Me.LBLmenu(Item).Enabled = Estado
  
End Sub
