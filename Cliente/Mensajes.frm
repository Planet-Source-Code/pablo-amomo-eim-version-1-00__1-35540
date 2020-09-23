VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Mensajes 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   Icon            =   "Mensajes.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Mensajes.frx":000C
   ScaleHeight     =   5445
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer EventosDeusuarios 
      Interval        =   65000
      Left            =   2520
      Top             =   2010
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6940
      MouseIcon       =   "Mensajes.frx":38BD
      MousePointer    =   99  'Custom
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   51
      Top             =   4470
      Width           =   165
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6940
      MouseIcon       =   "Mensajes.frx":3A0F
      MousePointer    =   99  'Custom
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   50
      Top             =   4080
      Width           =   165
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6805
      Picture         =   "Mensajes.frx":3B61
      ScaleHeight     =   615
      ScaleWidth      =   315
      TabIndex        =   49
      Top             =   4050
      Width           =   310
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4980
      Top             =   1740
   End
   Begin VB.ListBox ListadoAmigosMultiChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   610
      Left            =   5460
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   4050
      Width           =   1640
   End
   Begin VB.Timer Animacion 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8400
      Top             =   1920
   End
   Begin RichTextLib.RichTextBox MensajeEnviar 
      Height          =   585
      Left            =   135
      TabIndex        =   0
      Top             =   3840
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   1032
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Mensajes.frx":3CBA
   End
   Begin RichTextLib.RichTextBox VentanaMensajes 
      Height          =   1845
      Left            =   150
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3254
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Mensajes.frx":3D3C
      MouseIcon       =   "Mensajes.frx":3DB3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   5370
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Image Image3 
      Height          =   165
      Left            =   180
      Picture         =   "Mensajes.frx":3F15
      Stretch         =   -1  'True
      Top             =   4470
      Width           =   165
   End
   Begin VB.Label BotonFuente 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Mensajes.frx":4321
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label TipoDeFuente 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1440
      TabIndex        =   48
      Top             =   3450
      Width           =   1185
   End
   Begin VB.Label BotonTamano 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2700
      MouseIcon       =   "Mensajes.frx":4473
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label EtiquetaTamano 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   2700
      TabIndex        =   46
      Top             =   3525
      Width           =   1035
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   2700
      TabIndex        =   45
      Top             =   3375
      Width           =   1035
   End
   Begin VB.Shape Shape34 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   16
      Left            =   6795
      TabIndex        =   44
      Top             =   3375
      Width           =   165
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   16
      Left            =   6780
      Top             =   3360
      Width           =   195
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   15
      Left            =   6615
      TabIndex        =   43
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   14
      Left            =   6435
      TabIndex        =   42
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   13
      Left            =   6255
      TabIndex        =   41
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   12
      Left            =   6075
      TabIndex        =   40
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   11
      Left            =   5895
      TabIndex        =   39
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   10
      Left            =   5715
      TabIndex        =   38
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   9
      Left            =   5535
      TabIndex        =   37
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   8
      Left            =   5355
      TabIndex        =   36
      Top             =   3555
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   7
      Left            =   6615
      TabIndex        =   35
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   6
      Left            =   6435
      TabIndex        =   34
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   5
      Left            =   6255
      TabIndex        =   33
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   4
      Left            =   6075
      TabIndex        =   32
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   3
      Left            =   5895
      TabIndex        =   31
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   2
      Left            =   5715
      TabIndex        =   30
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   1
      Left            =   5535
      TabIndex        =   29
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label BotonColores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   5355
      TabIndex        =   28
      Top             =   3375
      Width           =   165
   End
   Begin VB.Shape ShapeColores 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   15
      Left            =   6600
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   14
      Left            =   6420
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   13
      Left            =   6240
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   12
      Left            =   6060
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   11
      Left            =   5880
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   10
      Left            =   5700
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   9
      Left            =   5520
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   8
      Left            =   5340
      Top             =   3540
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   7
      Left            =   6600
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   6420
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   5
      Left            =   6240
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   4
      Left            =   6060
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   3
      Left            =   5880
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   5700
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   1
      Left            =   5520
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape ShapeColores 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   5340
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image ScrollAbajo 
      Height          =   240
      Left            =   6860
      MouseIcon       =   "Mensajes.frx":45C5
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image ScrollArriba 
      Height          =   240
      Left            =   6860
      MouseIcon       =   "Mensajes.frx":4717
      MousePointer    =   99  'Custom
      Top             =   1140
      Width           =   240
   End
   Begin VB.Label LabelMensajeEvento 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   390
      TabIndex        =   27
      Top             =   4485
      Width           =   4965
   End
   Begin VB.Label BotonCaras 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4230
      MouseIcon       =   "Mensajes.frx":4869
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   4230
      TabIndex        =   25
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Shape Shape29 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   4230
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label EstadoMultiChat 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   5505
      TabIndex        =   11
      Top             =   3850
      Width           =   1425
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   24
      Top             =   120
      Width           =   6045
   End
   Begin VB.Image IconoAplicacion 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   6570
      MouseIcon       =   "Mensajes.frx":49BB
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   6930
      MouseIcon       =   "Mensajes.frx":4B0D
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Image ImagenBloqueado 
      Height          =   240
      Left            =   2340
      MouseIcon       =   "Mensajes.frx":4C5F
      MousePointer    =   99  'Custom
      ToolTipText     =   "El Usuario se Encuentra Bloqueado..."
      Top             =   465
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label BotonVerBloqueos 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   1890
      MouseIcon       =   "Mensajes.frx":4DB1
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4890
      Width           =   1605
   End
   Begin VB.Label LabelVerBloqueos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2670
      TabIndex        =   21
      Top             =   4995
      Width           =   75
   End
   Begin VB.Shape Shape23 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   4870
      Width           =   1575
   End
   Begin VB.Label BotonNloquearAmigo 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   240
      MouseIcon       =   "Mensajes.frx":4F03
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4890
      Width           =   1605
   End
   Begin VB.Label LabelBloquearUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   4995
      Width           =   1575
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   250
      Shape           =   4  'Rounded Rectangle
      Top             =   4870
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   6855
      MouseIcon       =   "Mensajes.frx":5055
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Reenvia El Estatus de * MulTiChat *..."
      Top             =   800
      Width           =   240
   End
   Begin VB.Shape ShapeFuente 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label BotonUnderLine 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1020
      MouseIcon       =   "Mensajes.frx":51A7
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label LabelUnderLine 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   17
      Top             =   3450
      Width           =   285
   End
   Begin VB.Shape ShapeUnderline 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   1020
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label BotonItalic 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      MouseIcon       =   "Mensajes.frx":52F9
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label LabelItalic 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   15
      Top             =   3450
      Width           =   285
   End
   Begin VB.Shape ShapeItalic 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label BotonBold 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   180
      MouseIcon       =   "Mensajes.frx":544B
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label LabelBold 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   12
      Top             =   3450
      Width           =   285
   End
   Begin VB.Shape ShapeBold 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   375
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image AnimacionImagen 
      Height          =   255
      Left            =   4440
      Top             =   4890
      Width           =   270
   End
   Begin VB.Label BotonEnviarMensaje 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5400
      MouseIcon       =   "Mensajes.frx":559D
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4860
      Width           =   1605
   End
   Begin VB.Label LabelEnviarMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6150
      TabIndex        =   10
      Top             =   4995
      Width           =   75
   End
   Begin VB.Shape ShapeNo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   435
      Left            =   5415
      Shape           =   4  'Rounded Rectangle
      Top             =   4875
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2710
      TabIndex        =   8
      Top             =   495
      Width           =   1185
   End
   Begin VB.Label UltimaResepcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3915
      TabIndex        =   2
      Top             =   495
      Width           =   3195
   End
   Begin VB.Label EstadoUsuarioTexto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   450
      TabIndex        =   3
      Top             =   495
      Width           =   2115
   End
   Begin VB.Image EstadoUsuarioImagen 
      Height          =   240
      Left            =   145
      Top             =   470
      Width           =   240
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4530
      MouseIcon       =   "Mensajes.frx":56EF
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   820
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   6525
      MouseIcon       =   "Mensajes.frx":5841
      MousePointer    =   99  'Custom
      Top             =   800
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4155
      MouseIcon       =   "Mensajes.frx":5993
      MousePointer    =   99  'Custom
      Top             =   800
      Width           =   240
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2160
      MouseIcon       =   "Mensajes.frx":5AE5
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   820
      Width           =   1980
   End
   Begin VB.Image FlechaAbajo 
      Height          =   240
      Left            =   1805
      MouseIcon       =   "Mensajes.frx":5C37
      MousePointer    =   99  'Custom
      Top             =   800
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   150
      MouseIcon       =   "Mensajes.frx":5D89
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   820
      Width           =   1650
   End
End
Attribute VB_Name = "Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

' **************************************************************
' Detecta los Eventos de los Usuarios...
' **************************************************************
Public EventosDeUsuarioTexto As String
Public EnvioEvento As Boolean

' **************************************************************
' Variable para detectar el Control + Enter
' **************************************************************
Public ControlEnter As String

' **************************************************************
' Define la Respuesta recibida de cada uno de los Usuarios
' en el Multichat...
' **************************************************************
Private RespuestaUsuarioMultichat() As Integer

' **************************************************************
' Public Estado Anterior...
' **************************************************************
'Public EstadoAnteriorAmigoNumero As Integer
'Public EstadoAnteriorAmigoTexto As String

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

Private TiempoLogueoInicial As Date
Private Anterior As String
Private IndiceAnimacion As Integer

' **************************************************************
' Bandera para definir la primera vez que pasa por estado UnkNow....
' **************************************************************
Public PrimerEstado As Boolean
'Public UltimoEstadoOtros As String
'Public UltimoEstadoOtros As String

Private Type MiAmigoEnChat
 IDAmigoAlias As String
 EstadoNumerico As String
 Sexo As String
 Estadotexto As String
 ConfirmadoMultiChat As Boolean
 IDVentanaAmigo As String   ' Aca se debine el Handle de la Ventana del Amigo en su Cliente
                            ' "REMOTO"
 SeguiEnviando As Boolean
End Type

Private Type MiAmigoEnChatPendiente
 IDAmigoAlias As String
 Estado As Integer
 IDVentanaAmigo As String   ' Aca se debine el Handle de la Ventana del Amigo en su Cliente
                            ' "REMOTO"
End Type

' **************************************************************
' Variables Usadas para el Hipervinculo en el Rich
' **************************************************************
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI__
    X As Long
    Y As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private HiperVinculo As String

' **************************************************************
' Fecha y Hora y Usuario del Ultimo Mensaje
Public WithEvents MenuDesbloqueo As IcoMenu
Attribute MenuDesbloqueo.VB_VarHelpID = -1
Public WithEvents MenuDeEstadoDeUsuario As IcoMenu
Attribute MenuDeEstadoDeUsuario.VB_VarHelpID = -1
Public WithEvents MenuDeEleccionMultiChat As IcoMenu
Attribute MenuDeEleccionMultiChat.VB_VarHelpID = -1
Public WithEvents MenuDeEliminacionDeUsuario As IcoMenu
Attribute MenuDeEliminacionDeUsuario.VB_VarHelpID = -1
Public WithEvents MenuDeCaras As IcoMenu
Attribute MenuDeCaras.VB_VarHelpID = -1
Public WithEvents MenuDeTamanos As IcoMenu
Attribute MenuDeTamanos.VB_VarHelpID = -1
Public WithEvents MenuBLoqueo As IcoMenu
Attribute MenuBLoqueo.VB_VarHelpID = -1
Public UltimoMensajeHoraYFecha As String
Public UltimoMensajeUsuario As String
' Define si el Mensaje se recibio o no...
Public RecibidoOk As Integer
Public RecibidoOkRespuesta As String
' Con que usuarios se esta Chateando
Private AmigosEnChat() As MiAmigoEnChat
Private AmigosEnChatPendiente() As MiAmigoEnChatPendiente
' Cantidad de Usuarios con los Que se Chatea
Public CantidadDeAmigosEnChat As Integer
Public CantidadDeAmigosEnChatPendiente As Integer
' Define el Responsable del MultiCHAT
Public ResponsableMultiChat As Boolean
Public ResponsableMultichatUsuario As String
Public ResponsableMultichatVentanaID As String
Public Sub AgregarLineaGrisConDatos(IDAmigo As Integer, Optional Texto As String, Optional UsuarioAEnviar As String)
Dim Estado, VentanaID As String
Dim Contador As Integer
Dim RichTemporalX As Object

  ' *********************************************************************
  ' Crea el RicBox
  ' *********************************************************************
  Set RichTemporalX = CreateObject("RICHTEXT.RichtextCtrl.1")
   
  ' *********************************************************************
  ' Borra el Texto del Rich Temporal
  ' *********************************************************************
  'RichTemporalX.TextRTF = ""
  RichTemporalX.SelStart = 0
  
  ' *********************************************************************
  ' Carga la Linea Gris...
  ' *********************************************************************
  Clipboard.Clear
  Clipboard.SetData Cliente.LineaGris.Picture
  SendMessage RichTemporalX.hwnd, &H302, 0, 0
  Clipboard.Clear
  RichTemporalX.SelRTF = "{{\par }\b  }"
  
  ' *********************************************************************
  ' POne los Datos Correspondientes...
  ' *********************************************************************
  ' Pone la Imagen de Aviso...
  Clipboard.Clear
  Clipboard.SetData Cliente.ImagenAviso.Picture ' Craga la Imagen...
  SendMessage RichTemporalX.hwnd, &H302, 0, 0
  Clipboard.Clear
  ' Pone el Texto...
  If Trim(Texto) = "" Then
    Select Case CStr(Trim(AmigosEnChat(IDAmigo).EstadoNumerico))
     Case "-1" ' Usuario No Existe...
      Estado = MensajeRecurso(342)
     Case "0" ' 0. No Conectado
      Estado = MensajeRecurso(287)
     Case "1" ' 1. Visible Normal
      Estado = MensajeRecurso(180)
     Case "2" ' 2. No Disponible
      Estado = MensajeRecurso(181)
     Case "3" ' 3. Custom
      Estado = Varios.ArreglarLenguaje(Trim(AmigosEnChat(IDAmigo).Estadotexto))
      If Len(Estado) <= 3 Then
        Estado = Estado & "..."
       Else
        If Mid$(Estado, Len(Estado) - 2) <> "..." Then
         Estado = Estado & "..."
        End If
      End If
    End Select
    RichTemporalX.SelRTF = "{{{{\colortbl ;\red128\green128\blue128;}" & _
                             "{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & _
                             "\viewkind4\uc1\pard\cf1\b\fs16 " & _
                             "    " & AmigosEnChat(IDAmigo).IDAmigoAlias & " " & MensajeRecurso(286) & " " & Trim(Estado) & "\cf0\b0\f1\fs17}\par }\b  }"
   Else
    RichTemporalX.SelRTF = "{{{{\colortbl ;\red128\green128\blue128;}" & _
                             "{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & _
                             "\viewkind4\uc1\pard\cf1\b\fs16 " & _
                             "\b    " & Texto & "\cf0\b0\f1\fs17}\par }\b  }"
  End If
  
  ' *********************************************************************
  ' Si un Usuario se Desconecto y/o No Existe, y hay Multichat,...
  ' lo Saca...
  ' *********************************************************************
  If Me.CantidadDeAmigosEnChat > 1 Then
   If CStr(Trim(AmigosEnChat(IDAmigo).EstadoNumerico)) = "-1" Or CStr(Trim(AmigosEnChat(IDAmigo).EstadoNumerico)) = "0" Then 'Or CStr(Trim(AmigosEnChat(IDAmigo).EstadoNumerico)) = "2" Then
    For Contador = 1 To CantidadDeAmigosEnChat
     If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(IDAmigo)) Then
      VentanaID = Trim(AmigosEnChat(Contador).IDVentanaAmigo)
      ' Solo si es numerico...
      If IsNumeric(VentanaID) Then
       Varios.EnviarBorradousuario Me.ResponsableMultichatUsuario, Me.ResponsableMultichatVentanaID, Trim(IDAmigo), CStr(VentanaID)
      End If
     End If
    Next
   End If
  End If
  ' Si pasa a No Disponible, entonces, cambia la FLAG para que recarge el Listado...
  Me.CargarAmigosMultiChat
  
  ' *********************************************************************
  ' Carga la Linea Gris...
  ' *********************************************************************
  Clipboard.Clear
  Clipboard.SetData Cliente.LineaGris.Picture
  SendMessage RichTemporalX.hwnd, &H302, 0, 0
  Clipboard.Clear
  
  ' *********************************************************************
  ' Agrega Todo o envia a Usuario
  ' *********************************************************************
  If UsuarioAEnviar <> "" Then
   '' ACA ENVIA EL MENSAJE A????
  End If
  ' Lo Pone en la Ventana correspondienre...
  PonerTextoEnVentanaMensaje 0, "", RichTemporalX.TextRTF
       
End Sub
Public Sub CargarAmigosMultiChat()
Dim Contador As Integer
Dim Bloqueado, ResponsableDelMultichat, UsuarioNoDisponible As String
Dim UsuarioNoExiste, UsuarioNoConectado As String

 ' **************************************************************
 ' Primero Verifica que Exista MultiCHAT
 ' **************************************************************
 If CantidadDeAmigosEnChat = 1 Then
   Me.ListadoAmigosMultiChat.Clear
   ' * MultiChat * - No
   Me.EstadoMultiChat = MensajeRecurso(356)
   ' No Hay Amigos...
   Me.ListadoAmigosMultiChat.AddItem (MensajeRecurso(360))
   Me.LabelBloquearUsuario.Enabled = True
   If UsuarioBloqueado(Trim(AmigosEnChat(1).IDAmigoAlias)) Then
     Me.ImagenBloqueado.Visible = True
    Else
     Me.ImagenBloqueado.Visible = False
   End If
   Exit Sub
 End If
 
 ' Como es MultiChat esconde el Icono de Bloqueado...
 Me.ImagenBloqueado.Visible = False
 
 ' Como es Multichat Bloquea el Boton de Bloque de Amigo...
 Me.LabelBloquearUsuario.Enabled = False
 
 ' **************************************************************
 ' Carga los items de MultiChat
 ' **************************************************************
 Me.ListadoAmigosMultiChat.Clear
   ' * MultiChat * - Sí
   Me.EstadoMultiChat = MensajeRecurso(361)
 For Contador = 1 To CantidadDeAmigosEnChat
   If UsuarioBloqueado(Trim(AmigosEnChat(Contador).IDAmigoAlias)) Then
     '  [BLQ]
     Bloqueado = MensajeRecurso(362)
    Else
     Bloqueado = ""
   End If
   If Trim(AmigosEnChat(Contador).EstadoNumerico) = "-1" Then
     UsuarioNoExiste = "(-) " ' No Existe
    Else
     UsuarioNoExiste = " "
   End If
   If Trim(AmigosEnChat(Contador).EstadoNumerico) = "0" Then
     UsuarioNoConectado = "(!) " ' No Conectado
    Else
     UsuarioNoConectado = " "
   End If
   If Trim(AmigosEnChat(Contador).EstadoNumerico) = "2" Then
     UsuarioNoDisponible = "(X) " ' No Disponible
    Else
     UsuarioNoDisponible = " "
   End If
   If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(Me.ResponsableMultichatUsuario)) And CLng(UCase(Trim(AmigosEnChat(Contador).IDVentanaAmigo))) = CLng(UCase(Trim(Me.ResponsableMultichatVentanaID))) Then
     ResponsableDelMultichat = "(*) " ' Owner Multichat...
    Else
     ResponsableDelMultichat = " "
   End If
   Me.ListadoAmigosMultiChat.AddItem ("  " & Trim(AmigosEnChat(Contador).IDAmigoAlias)) & ResponsableDelMultichat & UsuarioNoExiste & UsuarioNoConectado & UsuarioNoDisponible & Bloqueado
 Next
 
End Sub
Private Sub ConvertirHipervinculo(Box As RichTextBox)
    Dim hypStart, Contador As Integer
    'Dim befor As String
    'Dim after As String
    Dim cuvantAddress, KeyWord As String
    Dim hypEnd As Integer
      
    Dim separator1 As String
    Dim separator2 As String
    
    ' **************************************************************
    ' Variables de Retencion de Estado
    ' **************************************************************
    Dim Color As Long
    Dim Underline As Boolean
    
    ' **************************************************************
    ' Guarda el Estado actual del RichText Box
    ' **************************************************************
    If Not IsNull(Box.SelColor) Then
     Color = Box.SelColor
    End If
    If Not IsNull(Box.SelUnderline) Then
     Underline = Box.SelUnderline
    End If
    
    For Contador = 1 To 3
    ' Define que tipo de String va a buscar
     Select Case Contador
      Case 1:
       KeyWord = "http:"
      Case 2:
      KeyWord = "www."
      Case 3:
       KeyWord = "FTP."
     End Select

       
     hypStart = Box.Find(KeyWord)
        
     While hypStart >= 0
       ' **************************************************************
       ' Define donde termina el Hipervinculo
       ' **************************************************************
       separator1 = InStr(hypStart + 1, Box.Text, vbCr)
       separator2 = InStr(hypStart + 1, Box.Text, Chr(32))
       hypEnd = separator1
       ' **************************************************************
       ' Define que separador Utiliza
       ' **************************************************************
       If separator1 < separator2 And separator1 > 0 Then
         hypEnd = separator1
        ElseIf separator2 < separator1 And separator2 > 0 Then
         hypEnd = separator2
        ElseIf separator2 = 0 Then
         hypEnd = separator1
        ElseIf separator1 = 0 And separator2 > 0 Then
         hypEnd = separator2
        ElseIf separator2 = 0 And separator1 > 0 Then
         hypEnd = separator1
       End If
       If separator1 = 0 And separator2 = 0 Then hypEnd = Len(Box.Text) + 1
       
       ' **************************************************************
       'cuvantAddress = Mid(Box.Text, hypStart + 1, hypEnd - hypStart - 1)
       cuvantAddress = Mid(Box.Text, hypStart + 1, hypEnd - hypStart - 1)
       ' **************************************************************
               
       ' **************************************************************
       ' Pone el Texto como Hipervinculo
       ' **************************************************************
       Box.SelStart = hypStart
       Box.Find cuvantAddress, hypStart
       Box.SelUnderline = True
       Box.SelColor = vbBlue
       Box.SelStart = hypStart + 1
       hypStart = Box.Find(KeyWord, hypStart + 1, Len(Box.Text), 2)
       
     Wend
    Next
    
    ' **************************************************************
    ' Restaura los Atributos del ReachTextBox
    ' **************************************************************
    Box.SelColor = Color
    Box.SelUnderline = Underline
    
    ' **************************************************************
    ' Se mueve al Final del BOX
    ' **************************************************************
    Box.SelStart = Len(Box.Text) + 1
    ' Pone el estado de los botones...
    DefinirEstadodeBotones
    
End Sub
Private Sub DefinirEstadodeBotones()
Dim Contador As Integer
Dim Bandera As Boolean

 ' **************************************************************
 ' Define el Boton de Tamaño
 ' **************************************************************
 If Me.MensajeEnviar.SelFontSize = 0 Or IsNull(Me.MensajeEnviar.SelFontSize) Then
   '  Puntos
   Me.EtiquetaTamano = Configuracion.FontEstandarTamano & MensajeRecurso(363)
  Else
   '  Puntos
   Me.EtiquetaTamano = Me.MensajeEnviar.SelFontSize & MensajeRecurso(363)
 End If
 
 ' **************************************************************
 ' Define el Boton de TipoFuente
 ' **************************************************************
 If Me.TipoDeFuente.Font <> Null Then
   Me.TipoDeFuente.Font = Me.MensajeEnviar.SelFontName
  Else
   Me.TipoDeFuente.Font = Configuracion.FontEstandarNombre
 End If
 
 ' **************************************************************
 ' Define el Estado del Boton BOLD
 ' **************************************************************
 If Me.MensajeEnviar.SelBold = False Then
    Me.ShapeBold.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeBold.BackColor = Variables.FontMensajeBotonApretadoFondo
 End If
 
 ' **************************************************************
 ' Define el Estado del Boton Italic
 ' **************************************************************
 If Me.MensajeEnviar.SelItalic = False Then
   Me.ShapeItalic.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeItalic.BackColor = Variables.FontMensajeBotonApretadoFondo
 End If

 ' **************************************************************
 ' Define el Estado del Boton Undelline
 ' **************************************************************
 If Me.MensajeEnviar.SelUnderline = False Then
   Me.ShapeUnderline.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeUnderline.BackColor = Variables.FontMensajeBotonApretadoFondo
 End If

 ' **************************************************************
 ' Define el Marco del Color correspondiente...
 ' **************************************************************
 ' Pone todos los marcos escondidos...
 For Contador = 0 To 16
  BotonColores(Contador).BorderStyle = 0
 Next
 ' Busca el Color correspondiente
 Bandera = False
 For Contador = 0 To 16
  If ShapeColores(Contador).FillColor = Me.MensajeEnviar.SelColor Then
   Bandera = True
   BotonColores(Contador).BorderStyle = 1
  End If
 Next
 ' Si no existe el Color entonces lo guarda en la casilla 16
 If Bandera = False Then
  If IsNull(Me.MensajeEnviar.SelColor) Then
    ShapeColores(16).FillColor = &HE0E0E0
   Else
    ShapeColores(16).FillColor = Me.MensajeEnviar.SelColor
  End If
  BotonColores(16).BorderStyle = 1
 End If
 
End Sub
Private Sub BotonBold_Click()

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeBold
 
 ' **************************************************************
 ' Pone y Saca el BOLD del Texto
 ' **************************************************************
 If Me.MensajeEnviar.SelBold = True Then
   Me.MensajeEnviar.SelBold = False
   Me.ShapeBold.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeBold.BackColor = Variables.FontMensajeBotonApretadoFondo
   Me.MensajeEnviar.SelBold = True
 End If

End Sub
Public Sub AlFinalDeVentanaMensaje(Box As RichTextBox)
  
  ' **************************************************************
  ' Se posiciona al Final de los Mensajes
  ' **************************************************************
  ''Box.SelRTF = Len(Box.TextRTF) + 1
  'Box.SelRTF = Len(Box.TextRTF) + 1
  'Box.TextRTF = "{\par }"
  Box.SelStart = Len(Box.TextRTF) ''+ 1
  Box.SelLength = 1
  
End Sub
Private Sub BotonColores_Click(Index As Integer)
Dim Contador As Integer

 ' **************************************************************
 ' Borra todos los marcos... Y pone el nuevo Marco
 ' **************************************************************
 For Contador = 0 To 15
  BotonColores(Contador).BorderStyle = 0
 Next
 BotonColores(Index).BorderStyle = 1
 
 ' **************************************************************
 ' Setea el Color
 ' **************************************************************
 Me.MensajeEnviar.SelColor = ShapeColores(Index).FillColor
 
 End Sub

Private Sub BotonTamano_Click()

  Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape34
  
  MenuDeTamanos.ShowMenu Me.BotonTamano.Left + Me.Left, Me.BotonTamano.Top + Me.Top + 340

End Sub

Public Sub CargarTextos()

 ' Ingles...
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 ' Botontes...
 Me.Shape22.BackColor = Variables.ShapesBackColor
 Me.Shape22.BorderColor = Variables.ShapesBorderColor
 Me.Shape23.BackColor = Variables.ShapesBackColor
 Me.Shape23.BorderColor = Variables.ShapesBorderColor
 Me.Shape29.BackColor = Variables.ShapesBackColor
 Me.Shape29.BorderColor = Variables.ShapesBorderColor
 Me.Shape34.BackColor = Variables.ShapesBackColor
 Me.Shape34.BorderColor = Variables.ShapesBorderColor
 Me.ShapeBold.BackColor = Variables.ShapesBackColor
 Me.ShapeBold.BorderColor = Variables.ShapesBorderColor
 Me.ShapeFuente.BackColor = Variables.ShapesBackColor
 Me.ShapeFuente.BorderColor = Variables.ShapesBorderColor
 Me.ShapeItalic.BackColor = Variables.ShapesBackColor
 Me.ShapeItalic.BorderColor = Variables.ShapesBorderColor
 Me.ShapeUnderline.BackColor = Variables.ShapesBackColor
 Me.ShapeUnderline.BorderColor = Variables.ShapesBorderColor
 Me.ShapeNo.BackColor = Variables.ShapesBackColor
 Me.ShapeNo.BorderColor = Variables.ShapesBorderColor
 
 ' **************************************************************
 ' Define los textos  del Formulario
 ' **************************************************************
 Me.LabelUnderLine.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.LabelItalic.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.LabelBold.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.EtiquetaTamano.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.Label1 = MensajeRecurso(346)                ' Usuario En Chat (0)...
 Me.Label2 = MensajeRecurso(347)                ' Agregar Amigos MultiChat...
 Me.Label3 = MensajeRecurso(348)                ' Eliminar Amigos MultiChat...
 Me.Label4 = MensajeRecurso(349)                ' Ultimo Mensaje
 Me.Label5.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.Label5 = MensajeRecurso(350)                ' Emociones
 Me.LabelMensajeEvento = MensajeRecurso(351)    ' [Control] + [Enter] Envia el Mensaje...
 Me.EventosDeUsuarioTexto = MensajeRecurso(351)
 Me.Label7.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.Label7 = MensajeRecurso(352)                ' Tamaño
 Me.UltimaResepcion = MensajeRecurso(353)       ' Status...
 Me.EstadoUsuarioTexto = MensajeRecurso(354)    ' Estado...
 Me.LabelVerBloqueos.ForeColor = Variables.FontBotonesColor
 Me.LabelVerBloqueos = MensajeRecurso(355)      ' Ver Bloqueos
 Me.EstadoMultiChat = MensajeRecurso(356)       ' * MultiChat * - No
 Me.LabelEnviarMensaje.ForeColor = Variables.FontBotonesColor
 Me.LabelEnviarMensaje = MensajeRecurso(218)    ' Enviar Mensaje...
 Me.TipoDeFuente.ForeColor = Variables.FontMensajeBotonColorLetra
 Me.TipoDeFuente = MensajeRecurso(358)          ' Fuente
 Me.LabelBloquearUsuario.ForeColor = Variables.FontBotonesColor
 Me.LabelBloquearUsuario = MensajeRecurso(304)  ' Bloquear Amigo...
 Me.Image6.Picture = Cliente.Imagenes.ListImages("MultiChat").Picture
 Me.Image6.ToolTipText = MensajeRecurso(443)
 Me.ImagenBloqueado.Picture = Cliente.Imagenes.ListImages("UsuarioBloqueado").Picture
 Me.ImagenBloqueado.ToolTipText = MensajeRecurso(444)
 Me.Image4.Picture = Cliente.Imagenes.ListImages("Minimizar").Picture
 Me.Image5.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.FlechaAbajo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoAzul").Picture
 Me.Image1.Picture = Cliente.ImagenesFlecha.ListImages("AbajoVerde").Picture
 Me.Image2.Picture = Cliente.ImagenesFlecha.ListImages("AbajoRoja").Picture
 Me.ScrollArriba.Picture = Cliente.ImagenesFlecha.ListImages("ArribaAzul").Picture
 Me.ScrollAbajo.Picture = Cliente.ImagenesFlecha.ListImages("AbajoAzul").Picture
 Me.Picture2.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaArriba").Picture
 Me.Picture3.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaAbajo").Picture
 Me.AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture

End Sub


Private Sub EnviarEvento(Evento As Integer, Optional Dato As String)
Dim Contador As Integer
Dim ComandoAdicional As String
 
 ' **************************************************************
 ' Enviar Evento a Todos los usuarios...
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  ComandoAdicional = ComandoAdicional & CompletarCadena(AmigosEnChat(Contador).IDAmigoAlias, 16, "D", " ") & CompletarCadena(AmigosEnChat(Contador).IDVentanaAmigo, 10, "I", "0") & Evento & Dato
 Next
 ComandoAdicional = "M" & CompletarCadena(CStr(CantidadDeAmigosEnChat), 2, "I", "0") & ComandoAdicional
 EnviarPaqueteTCP "43" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
                  ComandoAdicional
 
End Sub


Private Sub EventosDeusuarios_Timer()

 ' **************************************************************
 ' Una ves pasado un Minuto Elimiona el Evento...
 ' **************************************************************
 If EventosDeusuarios Then
  Me.AgregarEventoDelUsuario ""
 End If
 
End Sub

Private Sub Form_Resize()

 ' **************************************************************
 ' Detienen el Flasheo de la Ventana
 ' **************************************************************
 Me.Timer1.Enabled = False
         
End Sub
Private Sub Form_Unload(Cancel As Integer)

 ' **************************************************************
 ' Envia el Evento que Define que Salio del Multichat...
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = False Then
  EnviarEvento 2
 End If

 ' **************************************************************
 ' Si es el Owner envia la Cancelacion del * MultiChat *
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = True Then
  EnviarEvento 3
 End If

End Sub

Private Sub Menudesbloqueo_Click(ByVal Index As Long, Tag As String)
Dim Respuesta As Integer

 ' **************************************************************
 ' Ejecutar el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Desbloquea el Amigo
 ' **************************************************************
 Select Case Index
  Case 0:
   ' Desbloqueo el Amigo
   ' ¿Está Seguro que Desea Desbloquear al Amigo [ % ]?...
   Respuesta = MostrarMSGBox(MensajeRecurso(303) & Trim(AmigosEnChat(1).IDAmigoAlias) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then Exit Sub
   Varios.ProcesarUsuariosBloqueados "Sacar", Trim(AmigosEnChat(1).IDAmigoAlias)
 End Select
 
End Sub
Private Sub ImagenBloqueado_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Toma la Posicion actual del Cursos
 ' **************************************************************
 GetCursorPos Posicion
 
 ' **************************************************************
 ' Muestra el Menu
 ' **************************************************************
 MenuDesbloqueo.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY

End Sub

Private Sub MensajeEnviar_GotFocus()

 ' **************************************************************
 ' Cuando preciona CTRL + Enter borra el Rich Box, para que no
 ' deje un espacio y linea en blanco...
 ' **************************************************************
  If UCase(Me.ControlEnter) = "SI" Then
   Me.ControlEnter = ""
   BorraMensajeEnviar
   DefinirEstadodeBotones
  End If

End Sub

Private Sub MensajeEnviar_KeyPress(KeyAscii As Integer)
Dim Respuesta As Long

 ' **************************************************************
 ' Si Habia pasado a Enseguida Vuelvo, lo devuelve a Estado
 ' Disponible
 ' **************************************************************
 If Variables.PasoAEnseguidaVuelvo = True Then
  Variables.PasoAEnseguidaVuelvo = False
  ' Usted fue pasado a 'Enseguida Vulevo' por Inactividad... ¿Desea Volver a Estado Disponible?...
  Respuesta = MostrarMSGBox(MensajeRecurso(365), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
  ' Pregunta que quiere hacer
  If Respuesta = vbYes Then
   Cliente.MenuDeEstadosDeUsuario_Click 0, ""
   Me.MensajeEnviar.SetFocus
   ' Se posiciona al Final de Mensaje Eviar, para hacer que el usuario no se entere
   ' del cambio...
   AlFinalDeVentanaMensaje Me.MensajeEnviar
  End If
 End If
 
 ' **************************************************************
 ' Cuando preciona CTRL + Enter manda el Mensaje
 ' **************************************************************
 If KeyAscii = 10 Then
   Me.ControlEnter = "SI"
   BotonEnviarMensaje_Click
 End If
         
 If Me.EnvioEvento = False Then
  Me.EnvioEvento = True
  EnviarEvento "1"
 End If
 
End Sub

Private Sub MensajeEnviar_SelChange()

 ' **************************************************************
 ' Cambio del estado de los botones
 ' **************************************************************
 DefinirEstadodeBotones
 
End Sub

Private Sub MenuDeTamanos_Click(ByVal Index As Long, Tag As String)
    
  Me.MensajeEnviar.SelFontSize = 8 + Index * 2
  '  Puntos
  Me.EtiquetaTamano = Me.MensajeEnviar.SelFontSize & MensajeRecurso(363)

End Sub

Private Sub MenuDeCaras_Click(ByVal Index As Long, Tag As String)
    
  Me.MensajeEnviar.SelText = " "
  Clipboard.Clear
  Clipboard.SetData Cliente.ImagenCaras.ListImages(Index + 1).Picture
  SendMessage MensajeEnviar.hwnd, &H302, 0, 0
  Clipboard.Clear
  Me.MensajeEnviar.SelText = " "

End Sub


Private Sub BotonCaras_Click()

  Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape29
 
 MenuDeCaras.ShowMenu Me.BotonCaras.Left + Me.Left, Me.BotonCaras.Top + Me.Top + 340

End Sub
Public Sub CambiarLetraRemoto(Letras As String)

 ' **************************************************************
 ' Setea el Nuevo Font
 ' **************************************************************
 If Letras <> "" Then
  Me.MensajeEnviar.SelFontName = Letras
 End If
  
 ' **************************************************************
 ' Cambia los Botones devido al Cambio de Letra....
 ' **************************************************************
 DefinirEstadodeBotones

End Sub
Private Sub BotonFuente_Click()
Dim TiposDeLetras As New EleccionTipoDeLetra

 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeFuente
 
 ' **************************************************************
 ' Habre el Formulario de Letras
 ' **************************************************************
 Set TiposDeLetras = New EleccionTipoDeLetra
 With TiposDeLetras
  ' Le Define cual es la FontActual
  If IsNull(Me.MensajeEnviar.SelFontName) Then
    .MostrarFormulario Configuracion.FontEstandarNombre, Me.hwnd, False
   Else
    .MostrarFormulario Me.MensajeEnviar.SelFontName, Me.hwnd, False
   End If
  ' Pone la nueva Font si es "" es por que se cancelo
 End With
 
End Sub
Private Sub BotonItalic_Click()

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeItalic
  
 ' **************************************************************
 ' Pone y Saca el Italic del Texto
 ' **************************************************************
 If Me.MensajeEnviar.SelItalic = True Then
   Me.MensajeEnviar.SelItalic = False
   Me.ShapeItalic.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeItalic.BackColor = Variables.FontMensajeBotonApretadoFondo

   Me.MensajeEnviar.SelItalic = True
 End If
 
End Sub

Private Sub BotonNloquearAmigo_Click()
Dim Respuesta As Long

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape22

 ' **************************************************************
 ' Si esta en MultiChat no Puede Bloquear al Amigo
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Bloquea al Amigo
 ' **************************************************************
 ' ¿Está Seguro que Desea Bloquear al Amigo [ % ]?...
 Respuesta = MostrarMSGBox(MensajeRecurso(300) & Trim(AmigosEnChat(1).IDAmigoAlias) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
 If Respuesta = vbNo Then Exit Sub
 Varios.ProcesarUsuariosBloqueados "Agregar", Trim(AmigosEnChat(1).IDAmigoAlias)
 
End Sub

Private Sub BotonUnderLine_Click()

 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeUnderline

 ' **************************************************************
 ' Pone y Saca el UnderLine del Texto
 ' **************************************************************
 If Me.MensajeEnviar.SelUnderline = True Then
   Me.MensajeEnviar.SelUnderline = False
   Me.ShapeUnderline.BackColor = Variables.ShapesBackColor
  Else
   Me.ShapeUnderline.BackColor = Variables.FontMensajeBotonApretadoFondo
   Me.MensajeEnviar.SelUnderline = True
 End If


End Sub
Private Sub BotonVerBloqueos_Click()

 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.Shape23

 Load UsuariosBloqueados
 UsuariosBloqueados.Show
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0
  Exit Sub
 End If

End Sub
Private Sub Image4_Click()

 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 Me.WindowState = vbMinimized
 
End Sub

Private Sub Image5_Click()

 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"

 Unload Me
 
End Sub


Private Sub Image6_Click()

 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Verifica que Exista un MultiChat para enviar el Listado...
 ' **************************************************************
 If CantidadDeAmigosEnChat <= 0 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica si es el Dueño del MultiChat...
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = False Then
  ' Usted no es el Responsable del * MultiChat * solo el Responsable Puede Agregar Amigos...
  MostrarMSGBox MensajeRecurso(477), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Envia el Listado MultiChat
 ' **************************************************************
 Me.EnviarListadoDeMultiChat
  
End Sub
Private Sub ListadoAmigosMultiChat_Click()
Dim NombreSeleccionado As String
Dim Posicion As Variables.POINTAPI
Dim Posicion1 As Integer

 ' **************************************************************
 ' Verifica que sea MultiChat
 ' **************************************************************
 If CantidadDeAmigosEnChat <= 1 Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Toma la Posicion del Mouse
 ' **************************************************************
 GetCursorPos Posicion

 ' **************************************************************
 ' Define el Nombre Elegigo
 ' **************************************************************
 Posicion1 = InStr(1, Trim(Me.ListadoAmigosMultiChat.List(Me.ListadoAmigosMultiChat.ListIndex)), "[")
 If Posicion1 <> 0 Then
   NombreSeleccionado = Trim(Mid$(Trim(Me.ListadoAmigosMultiChat.List(Me.ListadoAmigosMultiChat.ListIndex)), 1, Posicion1 - 1))
  Else
   NombreSeleccionado = Trim(Me.ListadoAmigosMultiChat.List(Me.ListadoAmigosMultiChat.ListIndex))
 End If
 
 ' **************************************************************
 ' Muestra el Menu
 ' **************************************************************
 CargarMenuEleccionDeUsuario NombreSeleccionado
 Me.MenuBLoqueo.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY

End Sub
Private Sub MenuBLoqueo_Click(ByVal Index As Long, Tag As String)
Dim Respuesta As Long

 Select Case Index
  Case 2
   ' **************************************************************
   ' Procesa un Desbloqueo
   ' **************************************************************
   Respuesta = UsuarioBloqueado(Trim(Tag))
   If Not Respuesta Then
    ' El Amigo [ % ] no se encuentra Bloqueado...
    Respuesta = MostrarMSGBox(MensajeRecurso(134) & Trim(Tag) & MensajeRecurso(134), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
    Exit Sub
   End If
   ' ¿Está Seguro que Desea Desbloquear al Amigo [ % ]?...
   Respuesta = MostrarMSGBox(MensajeRecurso(303) & Trim(Tag) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then Exit Sub
   Varios.ProcesarUsuariosBloqueados "Sacar", Trim(Tag)
   CargarAmigosMultiChat
  Case 0
   ' **************************************************************
   ' Procesa un Bloqueo de Usuario
   ' **************************************************************
   ' ¿Está Seguro que Desea Bloquear al Amigo [ % ]?...
   Respuesta = MostrarMSGBox(MensajeRecurso(300) & Trim(Tag) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then Exit Sub
   Varios.ProcesarUsuariosBloqueados "Agregar", Trim(Tag)
   CargarAmigosMultiChat
  Case 1
   ' **************************************************************
   ' Procesa un Mensaje Privado
   ' **************************************************************
   Respuesta = UsuarioBloqueado(Trim(Tag))
   If Respuesta Then
    ' El Amigo [ % "] se encuentra Bloqueado por lo Que no puede Entablar un Mensaje Privado... Primero debe Desbloquearlo...
    Respuesta = MostrarMSGBox(MensajeRecurso(134) & Trim(Tag) & MensajeRecurso(371), vbOKOnly, "vbInformation", Configuracion.TituloVentanas)
    Exit Sub
   End If
   
   Respuesta = BuscarVentana(Trim(Tag))
   ' Verifica si existe la Ventana si Existe, se
   ' posiciona en ella, sino abre una nueva...
   If Respuesta = 0 Then
     CrearVentanaMensaje Trim(Tag)
    Else
     ' Aca lo que se hace es que no tome como formulario
     ' el Actual, ya que por el Multichat la Busqueda resulta
     ' exitosa...
     If Forms(Respuesta).hwnd <> Me.hwnd Then
       Forms(Respuesta).Show
       Forms(Respuesta).SetFocus
      Else
       CrearVentanaMensaje Trim(Tag)
     End If
   End If

 End Select
 
End Sub
Private Sub MenuDeEliminacionDeUsuario_Click(ByVal Index As Long, Tag As String)
Dim VentanaID As String
Dim Contador, SegundosTranscurridos As Integer
Dim TiempoInicial As Date

 ' **************************************************************
 ' Verifica Que sea el Responsable del MultiCHAT
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = False Then
  ' Usted no es el Responsable del * MultiChat * solo el Responsable Puede Eliminar Amigos...
  MostrarMSGBox MensajeRecurso(134) & Trim(Me.ResponsableMultichatUsuario) & MensajeRecurso(480) & MensajeRecurso(372), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica que no se haya clickeado el No Hay Ususarios Dispo...
 ' **************************************************************
 If Trim(Me.MenuDeEliminacionDeUsuario.LBLmenu(Index).Tag) = "" Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Buscar la Ventana ID del Usuario que se esta borrando...
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(CStr(Trim(Me.MenuDeEliminacionDeUsuario.LBLmenu(Index).Tag))) Then
   VentanaID = CompletarCadena(CStr(Trim(AmigosEnChat(Contador).IDVentanaAmigo)), 10, "I", "0")
  End If
 Next
 
 ' **************************************************************
 ' Agregar un Usuario al MultiCHAT...
 ' **************************************************************
 TratarAmigosEnChat "Sacar", CStr(Trim(Me.MenuDeEliminacionDeUsuario.LBLmenu(Index).Tag))
 
 ' **************************************************************
 ' Envia el Listado actualizado de los Concurrentes al MultiChat
 ' **************************************************************
 Me.EnviarListadoDeMultiChat
 
 ' **************************************************************
 ' Espera 1 segundos...
 ' **************************************************************
 TiempoInicial = Time
 Do
 DoEvents
  If BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -1 And BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -2 Then Exit Do
  SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
  If SegundosTranscurridos >= 1 Then Exit Do
 Loop
   
 ' **************************************************************
 ' Al Usuario Borrador
 ' **************************************************************
 Me.EnviarListadoDeMultiChat CStr(Trim(Me.MenuDeEliminacionDeUsuario.LBLmenu(Index).Tag)), VentanaID
 
 ' **************************************************************
 ' Cuando no quedan Amigos, el Estado es no Responsable del
 ' MultiChat
 ' **************************************************************
 If CantidadDeAmigosEnChat = 1 Then
  ResponsableMultiChat = False
 End If
 
  

End Sub
Private Sub Image2_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + Image2.Left
 Posicion.Y = Me.Top + Image2.HeighT + Image2.Top + 5
 Me.CargarMenuEstadosyEliminacion
 Me.MenuDeEliminacionDeUsuario.ShowMenu Posicion.X, Posicion.Y
 
End Sub
Private Sub Label3_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + Label3.Left
 Posicion.Y = Me.Top + Label3.HeighT + Label3.Top + 5
 Me.CargarMenuEstadosyEliminacion
 Me.MenuDeEliminacionDeUsuario.ShowMenu Posicion.X, Posicion.Y

End Sub

' **************************************************************

Private Sub MenuDeEleccionMultiChat_Click(ByVal Index As Long, Tag As String)
Dim AmigosMultiChat, Respuesta2 As String
Dim Contador, SegundosTranscurridos, Respuesta As Integer
Dim TiempoInicial As Date
Dim AlgunoNoAcepto As Boolean

 ' **************************************************************
 ' Verifica Que sea el Responsable del MultiCHAT
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = False Then
  ' Usted no es el Responsable del * MultiChat * solo el Responsable Puede Agregar Amigos...
  MostrarMSGBox MensajeRecurso(134) & Trim(Me.ResponsableMultichatUsuario) & MensajeRecurso(480) & MensajeRecurso(367), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verifica Que no se haya pasado el Maximo de 10 Usuarios en el
 ' Chat...
 ' **************************************************************
 If CantidadDeAmigosEnChat = MaximoMultichat Then
  ' Ya Alcanzó el Maximo de Usuarios Concurrentes en MultiCHAT (10 Usuarios...)
  MostrarMSGBox MensajeRecurso(374), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Exit Sub
 End If

 ' **************************************************************
 ' Verifica que no se haya clickeado el No Hay Ususarios Dispo...
 ' **************************************************************
 If Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag) = "" Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Primero Verifica que el Usuario con el Cual se esta chatenado
 ' acepte el MultiChat... Si no Acepta, lo que hace es abrir un
 ' Chat normal contra el Usuario del MultiChat...
 ' **************************************************************
 ' Esto lo hace ya que sino, llega al Multichat (Cuando se agregan mas de 2 usuarios)
 ' y cancela por que respuesta=x
 Respuesta = 1
 AlgunoNoAcepto = False
 If CantidadDeAmigosEnChat = 1 Then
  ' **************************************************************
  ' Dispara la Animacion
  ' **************************************************************
  EstadoAnimacion True
  
  ' **************************************************************
  ' Anexa el Comando para Solicitud de Unirse a Multichat
  ' El Envio lo hace en referencia a los Dos usuarios que van a
  ' participar, el mismo, y el nuevo usuario que se esta
  ' incorporando via la seleccion en el Menu...
  ' **************************************************************
  AmigosMultiChat = "2" & "2" & CompletarCadena(Trim(Configuracion.IDAliasUsuario), 16, "D", " ") & CompletarCadena(CStr(Me.hwnd), 10, "I", "0") & CompletarCadena(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag), 16, "D", " ") & CompletarCadena("0", 10, "I", "0")
  ' **************************************************************
  ' Envia el Paquete
  ' **************************************************************
  EnviarPaqueteTCP "3" & CompletarCadena(Trim(AmigosEnChat(1).IDAmigoAlias), 16, "D", " ") & AmigosMultiChat
  
  ' **************************************************************
  ' Agregar el Usuario Actual, como un MultiChat
  ' **************************************************************
  Me.AgregarChayMultiusuarioPendiente "Agregar", Trim(AmigosEnChat(1).IDAmigoAlias)
  ' **************************************************************
  ' Espera 5 segundos por el OK del Usuario
  ' **************************************************************
  TiempoInicial = Time
  Do
  DoEvents
   If BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -1 And BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -2 Then Exit Do
   SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
   If SegundosTranscurridos >= Configuracion.TimeOutMultiChat Then Exit Do
  Loop
  ' **************************************************************
  ' Frena la Animacion
  ' **************************************************************
  EstadoAnimacion False
  
  ' **************************************************************
  ' Confirma si el Usuario Acepto, No Acepto, o No Contesto la
  ' solicitud....
  ' **************************************************************
  Respuesta = Me.BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias))
  Select Case Respuesta
   Case -1
    ' No se recibio Respuesta del Usuario [ & ] para unirse a su * MultiChat *...
    MostrarMSGBox MensajeRecurso(375) & Trim(AmigosEnChat(1).IDAmigoAlias) & MensajeRecurso(376), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    ' **************************************************************
    ' Elimina el Usuario como Pendiente de COnfirmacion...
    ' **************************************************************
    Me.AgregarChayMultiusuarioPendiente "Borrar", Trim(AmigosEnChat(1).IDAmigoAlias)
    AlgunoNoAcepto = True ' Bandera para reenviar el Listado de Multichat...
   Case 0
    ' El Usuario [ % ] no acepto su Solicitud de * MultiChat *...
    MostrarMSGBox MensajeRecurso(377) & Trim(AmigosEnChat(1).IDAmigoAlias) & MensajeRecurso(378), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    ' **************************************************************
    ' Elimina el Usuario como Pendiente de COnfirmacion...
    ' **************************************************************
    Me.AgregarChayMultiusuarioPendiente "Borrar", Trim(AmigosEnChat(1).IDAmigoAlias)
    AlgunoNoAcepto = True ' Bandera para reenviar el Listado de Multichat...
  End Select
 End If
 
 ' **************************************************************
 ' Confirma si el Usuario Acepto, No Acepto, o No Contesto
 ' cambiando el Ventana ID...
 ' **************************************************************
 TratarAmigosEnChat "AgregarPendiente", CStr(Trim(AmigosEnChat(1).IDAmigoAlias))
 ' **************************************************************
 ' Si el Usuario con el cual se esta chateando Acepta, entonces,
 ' Sigue creando el MultiChat sino crea una ventana contra el
 ' usuario que se pedia MultiChat
 ' **************************************************************
   
 If Respuesta <> 1 Then
  ' Como el usuario no acepto el Multichat, se pide si se quiere abrir una
  ' ventana de Mensaje contra el usuario al cual se queria incorporar al MultiChat...
  ' ¿Desea Abrir una Ventana de Mensajes con el Usuario [ % ]?...
  Respuesta = MostrarMSGBox(MensajeRecurso(379) & Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag) & MensajeRecurso(301), vbYesNo, "vbQuestion", Configuracion.TituloVentanas, False)
  If Respuesta = vbYes Then
   ComandosIntercambioDePaquete.ProcesarMensaje Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag), "U", "SI"
  End If
  Exit Sub
 End If
  
 ' **************************************************************
 ' Se Informa al Usuario que se lo Esta Tratando de Incluir en
 ' un Multichat...
 ' **************************************************************
 AmigosMultiChat = ""
 For Contador = 1 To CantidadDeAmigosEnChat
  AmigosMultiChat = AmigosMultiChat & CompletarCadena(AmigosEnChat(Contador).IDAmigoAlias, 16, "D", " ") & CompletarCadena(AmigosEnChat(Contador).IDVentanaAmigo, 10, "I", "0")
 Next
 ' Se agrega a el Mismo como Amigo para tener la referencia de la Ventana...
 Dim CantidadTemp As Integer
 AmigosMultiChat = CompletarCadena(Trim(Configuracion.IDAliasUsuario), 16, "D", " ") & CompletarCadena(Me.hwnd, 10, "I", "0") & AmigosMultiChat
 CantidadTemp = CantidadDeAmigosEnChat + 1
 If CantidadTemp = 10 Then
   AmigosMultiChat = "0" & AmigosMultiChat
  Else
   AmigosMultiChat = CantidadTemp & AmigosMultiChat
 End If
 
 ' **************************************************************
 ' Dispara la Animacion
 ' **************************************************************
 EstadoAnimacion True

 ' **************************************************************
 ' Anexa el Comando para Solicitud de Unirse a Multichat
 ' **************************************************************
 AmigosMultiChat = "2" & AmigosMultiChat
 ' **************************************************************
 ' Envia el Paquete
 ' **************************************************************
 EnviarPaqueteTCP "3" & CompletarCadena(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag), 16, "D", " ") & AmigosMultiChat
 
 ' **************************************************************
 ' Agrega el Amigo como Multichat Pendiente...
 ' **************************************************************
 Me.AgregarChayMultiusuarioPendiente "Agregar", Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)
 
 ' **************************************************************
 ' Espera 5 segundos por el OK del Usuario
 ' **************************************************************
 TiempoInicial = Time
 Do
 DoEvents
  If BuscarUsuarioPendiente(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)) <> -1 And BuscarUsuarioPendiente(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)) <> -2 Then Exit Do
  SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
  If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
 Loop
 
  ' **************************************************************
  ' frena la Animacion
  ' **************************************************************
  EstadoAnimacion False
 
 ' **************************************************************
 ' Confirma si el Usuario Acepto, No Acepto, o No Contesto la
 ' solicitud....
 ' **************************************************************
 Respuesta = Me.BuscarUsuarioPendiente(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag))
 Respuesta2 = Me.BuscarUsuarioPendienteHWND(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag))
 Select Case Respuesta
  Case -1
   ' No se recibio Respuesta del Usuario [ % ] para unirse a su * MultiChat *...
   MostrarMSGBox MensajeRecurso(375) & Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag) & MensajeRecurso(376), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   ' **************************************************************
   ' Elimina el Usuario como Pendiente de COnfirmacion...
   ' **************************************************************
   Me.AgregarChayMultiusuarioPendiente "Borrar", Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)
   AlgunoNoAcepto = True ' Bandera para reenviar el Listado de Multichat...
  Case 0
   MostrarMSGBox "El Usuario [" & Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag) & "] no acepto su Solicitud de * MultiChat *...", vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   ' **************************************************************
   ' Elimina el Usuario como Pendiente de COnfirmacion...
   ' **************************************************************
   Me.AgregarChayMultiusuarioPendiente "Borrar", Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)
   AlgunoNoAcepto = True ' Bandera para reenviar el Listado de Multichat...
 End Select
 
 ' **************************************************************
 ' Si alguno no acepto manda el Listado a todos los participantes
 ' **************************************************************
 If AlgunoNoAcepto Then
  ' **************************************************************
  ' Enviar a Todos los Usuarios el Nuevo Listado de MultiChat
  ' **************************************************************
  EnviarListadoDeMultiChat
 End If
 
 ' **************************************************************
 ' Elimina el Usuario como Pendiente de COnfirmacion...
 ' **************************************************************
 If Respuesta = -1 Or Respuesta = 0 Then
  Exit Sub
 End If

 ' **************************************************************
 ' Agregar un Usuario al MultiCHAT...
 ' **************************************************************
 TratarAmigosEnChat "Agregar", CStr(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)), BuscarUsuarioPendienteHWND(CStr(Trim(MenuDeEleccionMultiChat.LBLmenu(Index).Tag)))
 ResponsableMultiChat = True
 ' **************************************************************
 ' Enviar a Todos los Usuarios el Nuevo Listado de MultiChat
 ' **************************************************************
 EnviarListadoDeMultiChat
 
End Sub
'Public Sub EnviarListadoDeMultiChat(Optional Nombre As String, Optional VentanaID As String)
'Dim Contador As Integer
'Dim AmigosMultiChat  As String
'Dim PaqueteEnviar As String
'Dim ComandoAdicional As String
'
' ' **************************************************************
' ' Dispara la Animacion
' ' **************************************************************
'  EstadoAnimacion True
'
' ' **************************************************************
' ' Genera el Listado de Amigos
' ' **************************************************************
' AmigosMultiChat = ""
' For Contador = 1 To CantidadDeAmigosEnChat
'  AmigosMultiChat = AmigosMultiChat & CompletarCadena(AmigosEnChat(Contador).IDAmigoAlias, 16, "D", " ") & CompletarCadena(AmigosEnChat(Contador).IDVentanaAmigo, 10, "I", "0")
' Next
' ' **************************************************************
' ' Se agrega a el Mismo como Amigo del MultiChat
' ' **************************************************************
' Dim CantidadTemp As Integer
' AmigosMultiChat = AmigosMultiChat & CompletarCadena(Trim(Configuracion.IDAliasUsuario), 16, "D", " ") & CompletarCadena(Me.hwnd, 10, "I", "0")
' CantidadTemp = CantidadDeAmigosEnChat + 1
' If CantidadTemp = 10 Then
'   AmigosMultiChat = "0" & AmigosMultiChat
'  Else
'   AmigosMultiChat = CantidadTemp & AmigosMultiChat
' End If
'
' ' **************************************************************
' ' Si va a un nombre especifico lo envio a dicho Amigos
' ' **************************************************************
' If Trim(Nombre) <> "" Then
'  PaqueteEnviar = "3" & CompletarCadena(Trim(Nombre), 16, "D", " ") & "4" & VentanaID & AmigosMultiChat
'  ' **************************************************************
'  ' Envia el Paquete
'  ' **************************************************************
'  EnviarPaqueteTCP PaqueteEnviar
'  EstadoAnimacion False
'  Exit Sub
' End If
'
' ' **************************************************************
' ' Envia el Paquete a TODOS los Amigos...
' ' **************************************************************
' For Contador = 1 To Me.CantidadDeAmigosEnChat
'
'  ' **************************************************************
'  ' Anexa el Comando para Envio de Listado y la Ventana ID
'  ' **************************************************************
'  PaqueteEnviar = "3" & CompletarCadena(Trim(AmigosEnChat(Contador).IDAmigoAlias), 16, "D", " ") & "4" & CompletarCadena(AmigosEnChat(Contador).IDVentanaAmigo, 10, "I", "0") & AmigosMultiChat
'
'  ' **************************************************************
'  ' Espera para enviar el Paquete...
'  ' **************************************************************
'  Dim TiempoInicial As Date
'  Dim SegundosTranscurridos As Integer
'  TiempoInicial = Time
'  Do
'  DoEvents
'   If BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -1 And BuscarUsuarioPendiente(Trim(AmigosEnChat(1).IDAmigoAlias)) <> -2 Then Exit Do
'   SegundosTranscurridos = DateDiff("s", TiempoInicial, Time)
'   If SegundosTranscurridos >= 1 Then Exit Do
'  Loop
'
'  ' **************************************************************
'  ' Envia el Paquete
'  ' **************************************************************
'  EnviarPaqueteTCP PaqueteEnviar
'
' Next
'
' ' **************************************************************
' ' Frena la Animacion
' ' **************************************************************
' EstadoAnimacion False
'
'
'End Sub
Public Function BuscarUsuarioPendienteHWND(Nombre As String) As String
Dim Contador As Integer

 For Contador = 1 To CantidadDeAmigosEnChatPendiente
  If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
   BuscarUsuarioPendienteHWND = AmigosEnChatPendiente(Contador).IDVentanaAmigo
   Exit Function
  End If
 Next
  
End Function

Public Function BuscarUsuarioPendiente(Nombre As String) As Integer
Dim Contador As Integer

 ' **************************************************************
 ' -2 Usuario no existe en el Listado
 ' -1 No recibio respuesta
 '  0 No Quiere Participar
 '  1 Ok
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChatPendiente
  If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
   BuscarUsuarioPendiente = AmigosEnChatPendiente(Contador).Estado
   Exit Function
  End If
 Next
  
 BuscarUsuarioPendiente = -2
 
End Function
Private Sub PonerLabels()
Dim Usuario As String

   ' **************************************************************
   ' Define el Usuario
   ' **************************************************************
   Usuario = Trim(Configuracion.IDAliasUsuario)
   
   ' **************************************************************
   ' Pone el Titulo de la Ventana
   ' **************************************************************
   If CantidadDeAmigosEnChat = 1 Then
     '  - Charlando con
     Me.TituloVentana1 = Usuario & MensajeRecurso(382) & AmigosEnChat(1).IDAmigoAlias & "..."
     '  - Charlando con
     Me.Caption = Usuario & MensajeRecurso(382) & AmigosEnChat(1).IDAmigoAlias & "..."
     ' Usuario En Chat (1)...
     Me.Label1 = MensajeRecurso(383)
    Else
     '  - Charlando con *MultiChat*...
     Me.TituloVentana1 = Usuario & MensajeRecurso(384)
     '  - Charlando con *MultiChat*...
     Me.Caption = Usuario & MensajeRecurso(384)
     ' Usuario En Chat (
     Me.Label1 = MensajeRecurso(385) & CantidadDeAmigosEnChat & ")..."
   End If

End Sub
Public Function AgregarChayMultiusuarioPendiente(Metodo As String, Nombre As String, Optional Valor As Integer, Optional VentanaID As String)
Dim Contador As Integer

 Select Case Metodo
  ' **************************************************************
  ' Agrega un Amigo Pendiente de Confirmacion
  ' **************************************************************
  Case "Agregar"
   CantidadDeAmigosEnChatPendiente = CantidadDeAmigosEnChatPendiente + 1
   ReDim Preserve AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente)
   AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente).IDAmigoAlias = Nombre
   AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente).Estado = -1
   AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente).IDVentanaAmigo = ""
   
  ' **************************************************************
  ' Confirma
  ' **************************************************************
  Case "Confirmar"
   If CantidadDeAmigosEnChatPendiente = 0 Then Exit Function
   ' Busca el Registro a Modificas
   For Contador = 1 To CantidadDeAmigosEnChatPendiente
    ' Este es el Registro que debe Modificar
    If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
     AmigosEnChatPendiente(Contador).Estado = Valor
     AmigosEnChatPendiente(Contador).IDVentanaAmigo = VentanaID
    End If
   Next
   
   ' **************************************************************
   ' Borra
   ' **************************************************************
   Case "Borrar"
    If CantidadDeAmigosEnChatPendiente = 0 Then Exit Function
    ' Busca el Registro a Modificas
    For Contador = 1 To CantidadDeAmigosEnChatPendiente
     ' Este es el Registro que debe Modificar
     If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
      If Contador = CantidadDeAmigosEnChatPendiente Then
        CantidadDeAmigosEnChatPendiente = CantidadDeAmigosEnChatPendiente - 1
        Exit Function
       Else
        AmigosEnChatPendiente(Contador).IDAmigoAlias = AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente).IDAmigoAlias
        AmigosEnChatPendiente(Contador).Estado = AmigosEnChatPendiente(CantidadDeAmigosEnChatPendiente).Estado
        CantidadDeAmigosEnChatPendiente = CantidadDeAmigosEnChatPendiente - 1
         Exit Function
      End If
     End If
    Next
   
 End Select
 

End Function
Public Function TratarAmigosEnChat(Metodo As String, Nombre As String, Optional VentanaID As String) As Integer
 Dim Contador, Valor As Integer
 
 ' Metodo: Agregar, Buscar, Sacar, ModificarID
 Select Case Metodo
  ' **************************************************************
  ' Busca y un Ventana con el UsuarioID y el Handle
  ' **************************************************************
  Case "BuscarHandleYNombre"
   Valor = 0
   For Contador = 1 To CantidadDeAmigosEnChat
    'If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) And CLng(AmigosEnChat(Contador).IDVentanaAmigo) = CLng(VentanaID) Then
    If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) And CLng(AmigosEnChat(Contador).IDVentanaAmigo) = CLng(VentanaID) Then
     'AmigosEnChat(Contador).IDVentanaAmigo = VentanaID
     TratarAmigosEnChat = Contador
     Valor = Contador
     Exit For
    End If
   Next
   If Valor = 0 Then TratarAmigosEnChat = 0
  
  ' **************************************************************
  ' Cambi el ID de Ventana a Un Amigo Determinado
  ' **************************************************************
  Case "ModificarID"
   ' **************************************************************
   ' Busca el Amigo Pendiente y le Cambia el Ventana ID
   ' **************************************************************
   For Contador = 1 To CantidadDeAmigosEnChat
    If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
     AmigosEnChat(Contador).IDVentanaAmigo = VentanaID
    End If
   Next
     
  ' **************************************************************
  ' Pasa un Amigo Pendiente Ventana ID...
  ' **************************************************************
  Case "AgregarPendiente"
   ' **************************************************************
   ' Busca el Amigo Pendiente y le Cambia el Ventana ID
   ' **************************************************************
   For Contador = 1 To CantidadDeAmigosEnChatPendiente
    If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
     AmigosEnChat(1).IDVentanaAmigo = AmigosEnChatPendiente(Contador).IDVentanaAmigo
     ' **************************************************************
     ' Elimina el Usuario como Amigo Pendiente de MultiChat...
     ' **************************************************************
     Me.AgregarChayMultiusuarioPendiente "Borrar", Trim(Nombre)
     Exit For
    End If
   Next
        
  ' **************************************************************
  ' Agrega un Amigo al Listado de Amigos en CHAT
  ' **************************************************************
  Case "Agregar"
   ReDim Preserve AmigosEnChat(CantidadDeAmigosEnChat + 1)
   CantidadDeAmigosEnChat = CantidadDeAmigosEnChat + 1
   AmigosEnChat(CantidadDeAmigosEnChat).IDAmigoAlias = Nombre
   AmigosEnChat(CantidadDeAmigosEnChat).EstadoNumerico = "-1"
   AmigosEnChat(CantidadDeAmigosEnChat).Estadotexto = ""
   AmigosEnChat(CantidadDeAmigosEnChat).Sexo = "D"
   If Trim(VentanaID) <> "" Then
    AmigosEnChat(CantidadDeAmigosEnChat).IDVentanaAmigo = VentanaID
   End If
   If CantidadDeAmigosEnChat > 1 Then
    Me.ResponsableMultichatUsuario = Trim(Configuracion.IDAliasUsuario)
    Me.ResponsableMultichatVentanaID = Me.hwnd
   End If
   
   ' **************************************************************
   ' Busca el Amigo Pendiente y le Cambia el Ventana ID
   ' **************************************************************
   For Contador = 1 To CantidadDeAmigosEnChatPendiente
    If UCase(Trim(AmigosEnChatPendiente(Contador).IDAmigoAlias)) = UCase(Trim(Nombre)) Then
     AmigosEnChat(CantidadDeAmigosEnChat).IDVentanaAmigo = AmigosEnChatPendiente(Contador).IDVentanaAmigo
     ' **************************************************************
     ' Elimina el Usuario como Amigo Pendiente de MultiChat...
     ' **************************************************************
     AgregarChayMultiusuarioPendiente "Borrar", Trim(Nombre)
     Exit For
    End If
   Next
   
   ' **************************************************************
   ' Pone el Titulo de la Ventana
   ' **************************************************************
   PonerLabels
   
   ' **************************************************************
   ' Pone el Estado del Usuario o Multichat...
   ' **************************************************************
   MostrarLosEstados
   
   ' **************************************************************
   ' Carga el Estado del Nuevo Amigo
   ' **************************************************************
   PonerEstadoEnUsuario (CantidadDeAmigosEnChat)
   
   ' **************************************************************
   ' Carga el Listado de MultiChat
   ' **************************************************************
   CargarAmigosMultiChat
   ' **************************************************************
  
  ' **************************************************************
  ' Busca un Amigo en el Listao si esta devuelve
  ' 1 sino 0...
  ' **************************************************************
  Case "Buscar"
   For Contador = 1 To CantidadDeAmigosEnChat
    ' El Amigo esta
    If Trim(UCase(AmigosEnChat(Contador).IDAmigoAlias)) = Trim(UCase(Nombre)) Then
     TratarAmigosEnChat = 1
     Exit Function
    End If
   Next
   ' Si no esta devuelve un 0
   TratarAmigosEnChat = 0
  ' **************************************************************
  
  ' **************************************************************
  ' Saca un Amigo del Listado
  ' **************************************************************
  Case "Sacar"
   For Contador = 1 To CantidadDeAmigosEnChat
    If UCase(Trim(Nombre)) = Trim(UCase(AmigosEnChat(Contador).IDAmigoAlias)) Then
     Dim VentTEMP
     If CStr(VentanaID) = "" Then
       VentTEMP = "0"
      Else
       VentTEMP = CLng(VentanaID)
     End If
     If CStr(VentanaID) <> "" And CLng(VentTEMP) <> CLng(AmigosEnChat(Contador).IDVentanaAmigo) Then
       ' No hace nada ya que se pide se busque por VentanaId, pero la
       ' ventanaID no coincide...
      Else
       ' Si es el Ultimo directamente resta 1
       ' Sino copia el Ultimo a la Posicion Actual y resta 1
       If Contador = CantidadDeAmigosEnChat Then
         CantidadDeAmigosEnChat = CantidadDeAmigosEnChat - 1
         Exit For
        Else
         AmigosEnChat(Contador) = AmigosEnChat(CantidadDeAmigosEnChat)
         CantidadDeAmigosEnChat = CantidadDeAmigosEnChat - 1
         Exit For
       End If
     End If
    End If
   Next
   ' **************************************************************
   ' Pone el Titulo de la Ventana
   ' **************************************************************
   PonerLabels
   ' **************************************************************
   ' Pone el Estado del Usuario o Multichat...
   ' **************************************************************
   MostrarLosEstados
   ' **************************************************************
   ' Carga el Listado de MultiChat
   ' **************************************************************
   CargarAmigosMultiChat
  ' **************************************************************


End Select
 
 
End Function
Public Function BuscarEstadoNumerooTexto(Usuario As String, Variable As String) As String
Dim Contador As Integer

  ' **************************************************************
  ' Devolver EstadoNumerico o el Estado Texto...
  ' 1 sino 0...
  ' **************************************************************
  For Contador = 1 To CantidadDeAmigosEnChat
   ' El Amigo esta
   If Trim(UCase(AmigosEnChat(Contador).IDAmigoAlias)) = Trim(UCase(Usuario)) Then
    If UCase(Variable) = "NUMERO" Then
     BuscarEstadoNumerooTexto = CStr(AmigosEnChat(Contador).EstadoNumerico)
     Exit Function
    End If
    If UCase(Variable) = "TEXTO" Then
     BuscarEstadoNumerooTexto = CStr(AmigosEnChat(Contador).Estadotexto)
     Exit Function
    End If
   End If
  Next
   ' Si no esta devuelve un -1 como error
  BuscarEstadoNumerooTexto = "-1"
  ' **************************************************************

End Function
Private Sub Animacion_Timer()
 ' **************************************************************
 ' Timer que controla la animacion de la Conección
 ' **************************************************************
 ' Si el timer no es True sale
 If Animacion = False Then Exit Sub
 
 ' Verifica que figura debe mostrar
 IndiceAnimacion = IndiceAnimacion + 1
 If IndiceAnimacion = 5 Then IndiceAnimacion = 2
 
 AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(IndiceAnimacion).Picture

End Sub
Private Sub BotonEnviarMensaje_Click()
Dim Contador, Contador2, Contador3 As Integer
Dim Color, LargoDelTexto, Respuesta As Long
Dim LargoCalculadoPaquete, Valor1, Valor2 As Long
Dim Underline, Bloqueado As Boolean
Dim SegundosTranscurridos, UsuarioID As Integer
Dim ComandoAdicional, IDDelMensaje As String
Dim CantidadDePaquetesEnviar As Integer
Dim MensajePartido, MensajePartidoTemp, MensajeMultiUsuario As String
Dim Sacar() As String
Dim SacarTipo() As Integer
Dim SacarCantidad As Integer

 ' **************************************************************
 '      0 Enviado OK
 '      1 No Conectado
 '      2 No Disponible
 '      3 Usuario No Existe
 '     -1 Error En la Coneccion...
 ' **************************************************************
 
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
 
 ' **************************************************************
 ' Definir los Sacar a 0
 ' **************************************************************
 SacarCantidad = 0
 ReDim Sacar(Me.CantidadDeAmigosEnChat)
 ReDim SacarTipo(Me.CantidadDeAmigosEnChat)
 
 ' **************************************************************
 ' Define cuando se envio el Ultimo Mensaje
 ' **************************************************************
 Variables.UltimoMensajeEnviado = Time
 
 ' **************************************************************
 ' Efecto Boton
 ' **************************************************************
 EfectoBoton Me.ShapeNo
 
 ' **************************************************************
 ' Si el mensaje es Nulo, no manda nada...
 ' **************************************************************
 If Me.MensajeEnviar.Text = "" Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Graba los Atributos Actuales
 ' **************************************************************
 Color = MensajeEnviar.SelColor
 Underline = MensajeEnviar.SelUnderline
    
 ' **************************************************************
 ' Pasa a Hipervinculo
 ' **************************************************************
 ConvertirHipervinculo MensajeEnviar
 
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 ' Restaura los Atributos del ReachTextBox
 ' **************************************************************
 If IsNull(Color) Then
  Color = vbBlack
 End If
 If IsNull(Underline) Then
  Underline = False
 End If
 MensajeEnviar.SelColor = Color
 MensajeEnviar.SelUnderline = Underline
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 ' Si esta Desconectado no Manda Nada
 ' **************************************************************
 Select Case Configuracion.Logueado
  Case 0:
   ' En Estos Momentos no se Encuentra Conectado, Por tanto No Puede Enviar Mensajes...
   MostrarMSGBox MensajeRecurso(386), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Me.ControlEnter = ""
   MensajeEnviar.SetFocus
   Exit Sub
  Case 1:
   ' En Estos Momentos se Encuentra Conectandose al Sistema, Por tanto No Puede Enviar Mensajes...
   MostrarMSGBox MensajeRecurso(387), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   Me.ControlEnter = ""
   MensajeEnviar.SetFocus
   Exit Sub
 End Select
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 ' Verifica los Bloqueos...
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  ' **************************************************************
  ' Como esta bloqueado pregunta si desea desbloquearlo...
  ' **************************************************************
  If UsuarioBloqueado(AmigosEnChat(Contador).IDAmigoAlias) Then
   ' Tiene Bloqueado al usuario [ % ], por tanto no puede enviarle mensaje...¿Desea Desbloqueralo?
   Respuesta = MostrarMSGBox(MensajeRecurso(388) & Trim(AmigosEnChat(Contador).IDAmigoAlias) & MensajeRecurso(389), vbYesNo, "vbCritical", Configuracion.TituloVentanas)
   ' Lo Desbloquea
   If Respuesta = vbYes Then
    Varios.ProcesarUsuariosBloqueados "Sacar", Trim(AmigosEnChat(Contador).IDAmigoAlias)
    CargarAmigosMultiChat
   End If
   ' No lo Desbloquea y Avisa...
   If Respuesta = vbNo Then
    ' El Mensaje no fue Enviado...
    MostrarMSGBox MensajeRecurso(390), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   End If
  End If
 Next
 ' **************************************************************
 ' Si hay un solo amigo y esta bloqueado no manda nada...
 ' **************************************************************
 If CantidadDeAmigosEnChat = 1 And UsuarioBloqueado(AmigosEnChat(1).IDAmigoAlias) Then
  Exit Sub
 End If
 ' **************************************************************
 ' **************************************************************
 ' **************************************************************
 
 ' **************************************************************
 ' Define en Cuantas Partes debe enviar el Mensaje
 ' **************************************************************
 MensajePartidoTemp = MensajeEnviar.TextRTF
 CantidadDePaquetesEnviar = Int(Len(MensajePartidoTemp) / 4000)
 LargoDelTexto = CLng(Len(MensajePartidoTemp))
   
 ' **************************************************************
 ' Si la cantidad de Paquetes es mayor a 9 entonces
 ' cancela el envio (9*4000=36000)
 ' **************************************************************
 If LargoDelTexto > 12000 Then
  ' Avisa que "No es posible enviar un Mensaje tan Grande..."
  MostrarMSGBox MensajeRecurso(391), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
  Me.ControlEnter = ""
  Exit Sub
 End If
   
  ' **************************************************************
  ' Verifica si es la cantidad justa,o falta un paquete mas...
  ' Si Falta lo Agrega...
  ' **************************************************************
  LargoCalculadoPaquete = CantidadDePaquetesEnviar * CLng(4000)
  If LargoCalculadoPaquete <> LargoDelTexto Then
   CantidadDePaquetesEnviar = CantidadDePaquetesEnviar + 1
  End If
   
  ' **************************************************************
  ' Define que se le envien los Mensajes a todos los usuarios...
  ' Siempre que esten disponibles y/o Custom...
  ' **************************************************************
  Dim Bandera As Boolean
  Bandera = False
  For Contador = 1 To CantidadDeAmigosEnChat
   If AmigosEnChat(Contador).EstadoNumerico <> 0 And AmigosEnChat(Contador).EstadoNumerico <> 2 Then
    AmigosEnChat(Contador).SeguiEnviando = True
    Bandera = True
   End If
  Next
  If Bandera = False Then
   ' No Existen Usuario Disponibles o Conectados a los que se
   ' les pueda mandar Mensaje
   If CantidadDeAmigosEnChat > 1 Then
     MostrarMSGBox MensajeRecurso(472), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
    Else
     Select Case AmigosEnChat(1).EstadoNumerico
      Case 0:
       MostrarMSGBox MensajeRecurso(134) & AmigosEnChat(1).IDAmigoAlias & MensajeRecurso(176) & Chr$(13) & MensajeRecurso(473), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
      Case 2:
       MostrarMSGBox MensajeRecurso(134) & AmigosEnChat(1).IDAmigoAlias & MensajeRecurso(177) & Chr$(13) & MensajeRecurso(473), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
     End Select
   End If
   Exit Sub
  End If
  
  ' **************************************************************
  ' Dispara la Animacion...
  ' **************************************************************
  EstadoAnimacion True
   
  ' **************************************************************
  ' Pedir un ID para el mensaje
  ' **************************************************************
  IDDelMensaje = Varios.NuevoID
   
  ' **************************************************************
  ' Procesa el Envio por Partes
  ' **************************************************************
  For Contador2 = 1 To CantidadDePaquetesEnviar
   
     ' **************************************************************
     ' **************************************************************
     ' **************************************************************
     ' Si son Multiples Paquetes entre Paquete y Paquete espera...
     ' **************************************************************
     If CantidadDePaquetesEnviar > 1 Then
      ' **************************************************************
      ' Antes de Mandar espera un Segundo para evitar solapamientos...
      ' **************************************************************
      Dim TiempoLogueoInicial As Date
      ' **************************************************************
      ' Antes de Mandar espera un Segundo para evitar solapamientos...
      ' **************************************************************
      TiempoLogueoInicial = Time
      Do
       DoEvents
       SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
       If SegundosTranscurridos >= 1 Then Exit Do
      Loop
     End If
     ' **************************************************************
     ' **************************************************************
     ' **************************************************************
     
     ' **************************************************************
     ' **************************************************************
     ' **************************************************************
     ' En este punto Comando Adicional, define si el Mensaje es de Multichat o no
     ' para definirlo agrega un caracter antes del Mensaje. M=Multichat, U=Unico
     ' Si es Multichat,le anexa el HWnd de la Ventana de Destino...
     ' **************************************************************
     If CantidadDeAmigosEnChat > 1 Then
       ' **************************************************************
       ' Define el String de MultiUsuario
       ' **************************************************************
       Contador = 0
       For Contador3 = 1 To CantidadDeAmigosEnChat
        ' Solo si el usuario no esta bloqueado y la FLAG seguire enviando esta True
        If UsuarioBloqueado(AmigosEnChat(Contador3).IDAmigoAlias) = False And AmigosEnChat(Contador3).SeguiEnviando Then
         Contador = Contador + 1
         ComandoAdicional = ComandoAdicional & CompletarCadena(AmigosEnChat(Contador3).IDAmigoAlias, 16, "D", " ") & CompletarCadena(AmigosEnChat(Contador3).IDVentanaAmigo, 10, "I", "0")
        End If
       Next
       ComandoAdicional = "M" & CompletarCadena(CStr(CantidadDeAmigosEnChat), 2, "I", "0") & ComandoAdicional
       ' **************************************************************
       ' Si todos los usuarios estan bloqueado sale sin hacer nada...
       ' **************************************************************
       If Contador = 0 Then
        MostrarMSGBox MensajeRecurso(471), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
        Me.ControlEnter = ""
        Me.ControlEnter = ""
        EstadoAnimacion False
        MensajeEnviar.SelColor = Color
        MensajeEnviar.SelUnderline = Underline
        MensajeEnviar.SetFocus
        Exit Sub
       End If
      Else
       ComandoAdicional = "U"
     End If
     ' **************************************************************
     ' **************************************************************
     ' **************************************************************
     
     ' **************************************************************
     ' Prepara el Mensaje partido
     ' **************************************************************
     Valor1 = ((Contador2 - 1) * CLng(4000)) + 4000
     Valor2 = ((Contador2 - 1) * CLng(4000) + 1)
     If Valor1 > Len(MensajePartidoTemp) Then
       MensajePartido = Mid$(MensajePartidoTemp, Valor2)
      Else
       MensajePartido = Mid$(MensajePartidoTemp, Valor2, 4000)
     End If
     
     ' **************************************************************
     ' Si es Multichat, le pone el Handle de la Ventana pero despues
     ' de la parte n de n... Sino no funca
     ' **************************************************************
     If CantidadDeAmigosEnChat > 1 Then
       MensajePartido = CompletarCadena(AmigosEnChat(1).IDVentanaAmigo, 10, "I", "0") & MensajePartido
     End If
     
     ' **************************************************************
     ' Enviar el Mensaje Partido
     ' **************************************************************
     If Me.CantidadDeAmigosEnChat = 1 Then
       EnviarPaqueteTCP ("40" & CompletarCadena(AmigosEnChat(1).IDAmigoAlias, 16, "D", " ") & _
                          ComandoAdicional & _
                          CStr(Contador2) & CStr(CantidadDePaquetesEnviar) & CStr(IDDelMensaje) & MensajePartido)
      Else
       EnviarPaqueteTCP ("41" & CompletarCadena(AmigosEnChat(1).IDAmigoAlias, 16, "D", " ") & _
                           ComandoAdicional & _
                          CStr(Contador2) & CStr(CantidadDePaquetesEnviar) & CStr(IDDelMensaje) & MensajePartido)
     End If
   
     ' **************************************************************
     ' Espera 5 para validar el Envio del Mensaje
     ' **************************************************************
     TiempoLogueoInicial = Time
     RecibidoOk = -1
     Do Until RecibidoOk <> -1
      DoEvents
      SegundosTranscurridos = DateDiff("s", TiempoLogueoInicial, Time)
      If SegundosTranscurridos >= Configuracion.TimeOutGeneral Then Exit Do
     Loop
     
    ReDim RespuestaUsuarioMultichat(CantidadDeAmigosEnChat)
    ' Mensaje Mono-Usuario...
    If CantidadDeAmigosEnChat = 1 Then
     RespuestaUsuarioMultichat(1) = RecibidoOk
    End If
    ' Mensaje MultiUsuario...
    If CantidadDeAmigosEnChat > 1 Then
     ' **************************************************************
     ' Error de Coneccion...
     ' **************************************************************
     If RecibidoOk = -1 Then
      MostrarMSGBox Mid$(MensajeRecurso(396), 1, Len(MensajeRecurso(396)) - 1) & MensajeRecurso(400) & Mid$(MensajeRecurso(397), 2), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
      Me.ControlEnter = ""
      EstadoAnimacion False
      MensajeEnviar.SelColor = Color
      MensajeEnviar.SelUnderline = Underline
      If Me.Visible Then
       MensajeEnviar.SetFocus
      End If
      Me.EnvioEvento = False
      Exit Sub
     End If
     ' **************************************************************
     ' Error en el Paquete recibido - Largo invalido...
     ' **************************************************************
     If Len(RecibidoOkRespuesta) <> CantidadDeAmigosEnChat * 2 Then
      ' El Paquete es invalido...
      MostrarMSGBox MensajeRecurso(470), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
      ' Frena la Animacion y Sale
      Me.ControlEnter = ""
      EstadoAnimacion False
      MensajeEnviar.SelColor = Color
      MensajeEnviar.SelUnderline = Underline
      If Me.Visible Then
       MensajeEnviar.SetFocus
      End If
      Me.EnvioEvento = False
      Exit Sub
     End If
     ' **************************************************************
     ' Procesa la Respuesta Correspondiente...
     ' **************************************************************
     For Contador = 1 To CantidadDeAmigosEnChat
      If Not IsNumeric(Mid$(Trim(RecibidoOkRespuesta), Contador, 2)) Then
        RespuestaUsuarioMultichat(Contador) = -1
       Else
        RespuestaUsuarioMultichat(Contador) = Mid$(Trim(RecibidoOkRespuesta), 1 + (Contador - 1) * 2, 2)
      End If
     Next
     ' **************************************************************
     ' **************************************************************
    End If
    ' **************************************************************
    ' **************************************************************
    
    ' **************************************************************
    ' NOTA: Para que cuando se manda con Control Enter un Mensaje como la variable de
    ' ControlEnter con queda en SI, primero la pone en "" y la pasa a si si
    ' el mensaje salio OK. Esta variables es necesaria para que no deje espacios
    ' en Blanco...
    ' **************************************************************
    
    Dim PusoElMensaje As Boolean
    PusoElMensaje = False
    For UsuarioID = 1 To CantidadDeAmigosEnChat
     Select Case RespuestaUsuarioMultichat(UsuarioID)
      ' **************************************************************
      ' TODO OK...
      ' **************************************************************
      Case 0:
       ' Solo agrega el Mensaje a la Ventana del usuario, y borra
       ' el Mensaje cuando es el Ultimo (Previene Multiples Agregados
       ' en MultiChat
       If Contador2 = CantidadDePaquetesEnviar And PusoElMensaje = False Then
        ' Pone el Texto en la ventana Correspondiente...
        Me.PonerTextoEnVentanaMensaje vbBlack, Configuracion.IDAliasUsuario, Me.MensajeEnviar
        ' Borra el Ultimo mensaje escrito
        BorraMensajeEnviar
        ' Define como Bandera queya puso el Mensaje
        PusoElMensaje = True
       End If
      ' **************************************************************
      
      ' **************************************************************
      ' Usuario No Conectado...
      ' **************************************************************
      Case 1:
       ' Refresca a no Conectado...
       Varios.CambiarEstadoDeUsuario Trim(AmigosEnChat(UsuarioID).IDAmigoAlias), 0
       ' No fue Posible Enviar el Mensaje ya que [ % ] no Está Conectado...
       Respuesta = MostrarMSGBox(MensajeRecurso(392) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & MensajeRecurso(393) & MensajeRecurso(459), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
       If Respuesta = vbYes Then
        ' Primero verifica si existe un Mensaje Offline Abierto...
        Respuesta = Varios.BuscarVentanaMensajeOffLine(Trim(AmigosEnChat(UsuarioID).IDAmigoAlias))
        ' Si existe la muestra, Sino...
        If Respuesta <> 0 Then
          Forms(Respuesta).SetFocus
         Else
          CrearVentanaMensajeOffLine Trim(AmigosEnChat(UsuarioID).IDAmigoAlias), MensajeRecurso(454) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & "]..."
        End If
       End If
       AmigosEnChat(UsuarioID).SeguiEnviando = False
       ' Agrega el Amigo para ser Eliminado del Multichat...
       SacarCantidad = SacarCantidad + 1
       Sacar(SacarCantidad) = UsuarioID
       'SacarTipo(SacarCantidad) = 0
       ' **************************************************************
      
      ' **************************************************************
      ' Usuario No Disponible...
      ' **************************************************************
      Case 2:
       ' Refresca a no Disponible...
       Varios.CambiarEstadoDeUsuario Trim(AmigosEnChat(UsuarioID).IDAmigoAlias), 2
       ' No fue Posible Enviar el Mensaje ya que [ % ] no se Encuentra Disponible...
       Respuesta = MostrarMSGBox(MensajeRecurso(392) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & MensajeRecurso(177) & MensajeRecurso(459), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
       If Respuesta = vbYes Then
        ' Primero verifica si existe un Mensaje Offline Abierto...
        Respuesta = Varios.BuscarVentanaMensajeOffLine(Trim(AmigosEnChat(UsuarioID).IDAmigoAlias))
        ' Si existe la muestra, Sino...
        If Respuesta <> 0 Then
          Forms(Respuesta).SetFocus
         Else
          CrearVentanaMensajeOffLine Trim(AmigosEnChat(UsuarioID).IDAmigoAlias), MensajeRecurso(454) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & "]..."
        End If
       End If
       ' Lo setea para no seguir enviando...
       AmigosEnChat(UsuarioID).SeguiEnviando = False
      ' **************************************************************
      
      ' **************************************************************
      ' Usuario No Existe....
      ' **************************************************************
      Case 3:
       ' Refresca a no Existe...
       Varios.CambiarEstadoDeUsuario Trim(AmigosEnChat(UsuarioID).IDAmigoAlias), 3
       ' No fue Posible Enviar el Mensaje ya que [ % ] no Existe...
       MostrarMSGBox MensajeRecurso(392) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & MensajeRecurso(124), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
       ' Lo setea para no seguir enviando...
       AmigosEnChat(UsuarioID).SeguiEnviando = False
      ' **************************************************************
       ' Agrega el Amigo para ser Eliminado del Multichat...
       SacarCantidad = SacarCantidad + 1
       Sacar(SacarCantidad) = UsuarioID
       'SacarTipo(SacarCantidad) = 0
      ' **************************************************************
      ' Usuario No Existe....
      ' **************************************************************
      Case -1:
       ' No fue Posible Enviar el Mensaje a [ % ] debido a Un Error de Conección con el Servidor...
       MostrarMSGBox MensajeRecurso(396) & Trim(AmigosEnChat(UsuarioID).IDAmigoAlias) & MensajeRecurso(397), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
       ' Lo setea para no seguir enviando...
       AmigosEnChat(UsuarioID).SeguiEnviando = False
      ' **************************************************************
     
     End Select
    
    Next
   Next
  
  'End If
 
  ' **************************************************************
  ' Frena la Animacion
  ' **************************************************************
  EstadoAnimacion False
 
  ' **************************************************************
  ' Restaura los Atributos del ReachTextBox
  ' **************************************************************
  MensajeEnviar.SelColor = Color
  MensajeEnviar.SelUnderline = Underline

  ' **************************************************************
  ' Pone foco en la Ventana del Mensaje Enviar
  ' **************************************************************
  If Me.Visible Then
   MensajeEnviar.SetFocus
  End If
  
  ' **************************************************************
  ' Cuando Vuelve a Escribir envia el Evento...
  ' **************************************************************
  EnvioEvento = False
 
  ' **************************************************************
  ' Verifica que haya que eliminar algun Amigo...
  ' **************************************************************
  If SacarCantidad = 0 Then Exit Sub
  Dim PaqueteEnviar As String
  For Contador = 1 To SacarCantidad
   ' Le Envia un Evento al Responsable del Multichat para que
   ' Elimine los Amigos con Problemas...
   ' **************************************************************
   ' Espera un Segundo...
   ' **************************************************************
   TiempoLogueoInicial = Time
   Do Until DateDiff("s", TiempoLogueoInicial, Time) >= 1
    DoEvents
   Loop
   
   
   ' **************************************************************
   ' Enviar Evento a Todos los usuarios...
   ' **************************************************************
   If Me.ResponsableMultiChat Then
    Me.ResponsableMultichatUsuario = Trim(Configuracion.IDAliasUsuario)
    Me.ResponsableMultichatVentanaID = Me.hwnd
   End If
   
   Varios.EnviarBorradousuario Me.ResponsableMultichatUsuario, Me.ResponsableMultichatVentanaID, AmigosEnChat(CInt(Sacar(Contador))).IDAmigoAlias, CStr(AmigosEnChat(CInt(Sacar(Contador))).IDVentanaAmigo)
   'ComandoAdicional = "43" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & _
   '                 "M" & "01" & _
   '                 CompletarCadena(Trim(Me.ResponsableMultichatUsuario), 16, "D", " ") & _
   '                 CompletarCadena(Trim(Me.ResponsableMultichatVentanaID), 10, "I", "0") & "5" & _
   '                 CompletarCadena(Trim(), 10, "I", "0") & CompletarCadena(Trim(), 16, "D", " ")
   'EnviarPaqueteTCP ComandoAdicional
   'End Function
   ' **************************************************************
   ' Evento...
   ' **************************************************************
   'EnviarEvento 5, CompletarCadena(Me.ResponsableMultichatVentanaID, 10, "I", "0") & _

      
    
   ' Me.TratarAmigosEnChat "Sacar", Sacar(Contador)
   ' Me.CargarAmigosMultiChat
   ' If Me.CantidadDeAmigosEnChat = 1 Then
   '  If SacarTipo(Contador) = "-1" Then MostrarMSGBox MensajeRecurso(134) & Trim(Sacar(Contador)) & MensajeRecurso(478), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   '  If SacarTipo(Contador) = "0" Then MostrarMSGBox MensajeRecurso(134) & Trim(Sacar(Contador)) & MensajeRecurso(479), vbOKOnly, "vbCritical", Configuracion.TituloVentanas
   '  If ResponsableMultiChat Then
   '   Me.EnviarListadoDeMultiChat
   '  End If
   'End If
  Next
  
End Sub
Sub BorraMensajeEnviar()
 
 
 ' **************************************************************
 ' Borra MensajeEnviar y Define el Estado de los botones
 ' **************************************************************
 Me.MensajeEnviar = ""
 Me.MensajeEnviar.SelStart = 0
 DefinirEstadodeBotones
 
End Sub
Public Sub PonerTextoEnVentanaMensaje(Color As Long, NombreUsuario As String, Mensaje As String)
Dim TextoTMP, Agregar As String
Dim Contador As Integer

  ' *********************************************************************
  ' Primero se mueve al Final de la Ventana...
  ' *********************************************************************
  AlFinalDeVentanaMensaje Me.VentanaMensajes
     
  ' *********************************************************************
  ' Pone la Linea de [Amigo]: (EN AZUL/ARIAL/BOLD/8)
  ' *********************************************************************
  If Color = vbBlue Then TextoTMP = "\red0\green0\blue255"
  If Color = vbBlack Then TextoTMP = "\red0\green0\blue0"
  If Trim(NombreUsuario) <> "" Then
   Me.VentanaMensajes.SelRTF = "{{{\colortbl ;" & TextoTMP & ";}" & _
                               "{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & _
                               "\viewkind4\uc1\pard\cf1\b\fs16 " & _
                               Trim(NombreUsuario) & ": " & "\cf0\b0\f1\fs17}}"
  End If
      
  ' *********************************************************************
  ' Agrega el Mensaje...
  ' *********************************************************************
  Me.VentanaMensajes.SelRTF = "{" & CStr(Mensaje) & "\b  }"
    
  ' *********************************************************************
  ' Se Mueve al Final para asegurarse que muestre todo...
  ' *********************************************************************
  AlFinalDeVentanaMensaje Me.VentanaMensajes
 
End Sub
Public Sub MostrarVentana()

 ' **************************************************************
 ' Muestra la Ventana, Cargando la Info Necesario
 ' **************************************************************
 DefinirUltimaRecepcion "", ""
 PonerLabels
 Me.Show
  
End Sub
Private Sub FlechaAbajo_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + FlechaAbajo.Left
 Posicion.Y = Me.Top + FlechaAbajo.HeighT + FlechaAbajo.Top + 5
 Me.CargarMenuEstadosyEliminacion
 Me.MenuDeEstadoDeUsuario.ShowMenu Posicion.X, Posicion.Y

End Sub

Private Sub Form_Load()

 ' **************************************************************
 ' Carga los Textos del Formulario...
 ' **************************************************************
 Me.EventosDeUsuarioTexto = MensajeRecurso(351)
 EnvioEvento = False
 
 ' **************************************************************
 ' Carga los Textos del Formulario...
 ' **************************************************************
 Me.CargarTextos
  
 ' **************************************************************
 ' Pone un espacion para que quede alineado...
 ' **************************************************************
 Me.VentanaMensajes.SelRTF = "{{{\colortbl ;\red0\green0\blue255;}" & _
                               "{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & _
                               "\viewkind4\uc1\pard\cf1\b\fs16 " & _
                               " " & "\cf0\b0\f1\fs17}}"
                               
 ' **************************************************************
 ' Define el Estado Anterior...
 ' **************************************************************
 'EstadoAnteriorAmigoNumero = -10
 'EstadoAnteriorAmigoTexto = ""

 ' **************************************************************
 ' Define que la Cantidad de Amigos es 0
 ' **************************************************************
 CantidadDeAmigosEnChat = 0
 CantidadDeAmigosEnChatPendiente = 0
 Me.FormularioNombre = "Mensajes"
 ' Por ahora el Usuario no es responsable del MultiChat
 ResponsableMultiChat = False
 ' Carga el Menu que Muestra las Caras
 CargarMenusVarios
  
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.IconoAplicacion.Picture = Cliente.IconoAplicacion.Picture
 Me.Icon = Cliente.Icon
  
 ' **************************************************************
 ' Define el Font
 ' **************************************************************
 Me.MensajeEnviar.SelFontName = Configuracion.FontEstandarNombre
 Me.MensajeEnviar.SelFontSize = Configuracion.FontEstandarTamano
 DefinirEstadodeBotones
 
End Sub
Public Sub AgregarMensaje(RecibidoDe As String, Mensaje As String, HoraYFecha As String)
Dim Contador As Integer

 ' **************************************************************
 ' Si el Mensaje es Nulo descarta el Mensaje
 ' **************************************************************
 If Trim(Mensaje) = "" Then
  Exit Sub
 End If
 
 ' **************************************************************
 ' Agrega el Mensaje Recibido
 ' **************************************************************
 Me.PonerTextoEnVentanaMensaje vbBlue, RecibidoDe, Mensaje
 
 ' **************************************************************
 ' Define la Ultima Fecha y Hora de Recepcion...
 ' **************************************************************
 DefinirUltimaRecepcion Trim(RecibidoDe), HoraYFecha
   
 ' **************************************************************
 ' Verifica si hay modales abiertos...
 ' **************************************************************
 Dim Bandera As Boolean
 Bandera = True
 For Contador = 0 To Forms.Count - 1
  'DoEvents
  If Forms(Contador).FormularioNombre = "MensajesBox" Then
   If Forms(Contador).Modal Then Bandera = False
   Exit For
  End If
  If Forms(Contador).FormularioNombre = "IngresoBox" Then
   If Forms(Contador).Modal Then Bandera = False
   Exit For
  End If
 Next
 
 ' **************************************************************
 ' Flashea o Muestra
 ' **************************************************************
 If Me.WindowState = 1 Then
   Me.Timer1.Enabled = True
  Else
   ' **************************************************************
   ' Verifica que no haya ningun Modal Abierto Caso en el cual
   ' Espera...
   ' **************************************************************
   If Bandera Then
    Me.Show
   End If
 End If
         
 If Bandera Then
  MensajeEnviar.SetFocus
 End If
 
End Sub
Public Sub DefinirUltimaRecepcion(Usuario As String, HoraYFecha As String)
Dim Posicion As Integer

 ' **************************************************************
 ' Define que Todavia no Se Recibio Nada y Sale...
 ' **************************************************************
 If HoraYFecha = "" Then
  ' No fue Recibido Ningún Mensaje...
  Me.UltimaResepcion = MensajeRecurso(398)
  Me.UltimoMensajeHoraYFecha = HoraYFecha
  Me.UltimoMensajeUsuario = Trim(Usuario)
  PonerElEstadoDelUsuario
  Exit Sub
 End If
 
 ' **************************************************************
 ' Pone el Usuario Hora y Fecha
 ' **************************************************************
 Posicion = InStr(HoraYFecha, "_")
 '  a las  % del
 Me.UltimaResepcion = Usuario & MensajeRecurso(199) & Left(HoraYFecha, Posicion - 1) & MensajeRecurso(399) & Mid$(HoraYFecha, Posicion + 1)
 Me.UltimoMensajeHoraYFecha = HoraYFecha
 Me.UltimoMensajeUsuario = Trim(Usuario)
 
 ' **************************************************************
 ' Define el Estado del Amigo
 ' **************************************************************
 PonerElEstadoDelUsuario

End Sub
Public Sub PonerElEstadoDelUsuario() ' (Optional PonerMensaje As String)
Dim Contador As Integer
Dim Estado As String
Dim MultiChat As Boolean
Dim Usuario As String

 ' **************************************************************
 ' Define el Usuario...
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  PonerEstadoEnUsuario (Contador)
 Next
 
 MostrarLosEstados
 
End Sub
Private Sub MostrarLosEstados()  ' (Optional PonerMensaje As String)
Dim Contador As Integer
Dim Imagen, Picture  As String
Dim Contador2 As Integer
Dim TextoTMP

 ' **************************************************************
 ' Chat con un solo usuario
 ' **************************************************************
 'Texto = Me.EstadoUsuarioTexto
 If CantidadDeAmigosEnChat = 1 Then
  ' **************************************************************
  Me.EstadoUsuarioTexto = Trim(AmigosEnChat(1).Estadotexto)
  Select Case AmigosEnChat(1).EstadoNumerico
   Case "-1"
    Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages("UsuarioNoExiste").Picture
    Me.EstadoUsuarioTexto = MensajeRecurso(342)
   Case "0", "1", "2", "3"
    If AmigosEnChat(1).Sexo <> "D" And AmigosEnChat(1).Sexo <> "" Then
      Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages(AmigosEnChat(1).Sexo & AmigosEnChat(1).EstadoNumerico).Picture
     Else
      Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages("Desconocido").Picture
    End If
    If AmigosEnChat(1).EstadoNumerico = "3" Then
     Dim LargoTMP As Integer
     LargoTMP = Len(Me.EstadoUsuarioTexto)
     If LargoTMP <= 3 Then
       Me.EstadoUsuarioTexto = Me.EstadoUsuarioTexto & "..."
      Else
       If Mid$(Me.EstadoUsuarioTexto, LargoTMP - 2) <> "..." Then
        Me.EstadoUsuarioTexto = Me.EstadoUsuarioTexto & "..."
       End If
     End If
    End If
  End Select
 End If
  
 ' **************************************************************
 ' Chat * MultiCHAT *
 ' **************************************************************
  If CantidadDeAmigosEnChat > 1 Then
   ' * MultiChat *
   Me.EstadoUsuarioTexto = MensajeRecurso(400)
   Me.EstadoUsuarioImagen.Picture = Cliente.ImagenesAmigos.ListImages("MultiChat").Picture
  End If
  
    
End Sub

Public Sub CargarMenusVarios()

 ' **************************************************************
 ' Menu de Desbloqueo
 ' **************************************************************
 Set MenuDesbloqueo = New IcoMenu
  With MenuDesbloqueo
   ' Desbloquear...
   .SetItem 0, MensajeRecurso(401), Cliente.Imagenes.ListImages("EstadoVisible").Picture
  End With
  
 ' **************************************************************
 ' Carga el Menu de Caras
 ' **************************************************************
 Set MenuDeCaras = New IcoMenu
  With MenuDeCaras
   ' Felicidad...
   .SetItem 0, MensajeRecurso(402), Cliente.ImagenCaras.ListImages(1).Picture
   ' Risa...
   .SetItem 1, MensajeRecurso(403), Cliente.ImagenCaras.ListImages(2).Picture
   ' Tristeza...
   .SetItem 2, MensajeRecurso(404), Cliente.ImagenCaras.ListImages(3).Picture
   ' Llanto...
   .SetItem 3, MensajeRecurso(405), Cliente.ImagenCaras.ListImages(4).Picture
   ' Enojo...
   .SetItem 4, MensajeRecurso(406), Cliente.ImagenCaras.ListImages(5).Picture
   ' Complicidad...
   .SetItem 5, MensajeRecurso(407), Cliente.ImagenCaras.ListImages(6).Picture
   ' En las Nubes...
   .SetItem 6, MensajeRecurso(408), Cliente.ImagenCaras.ListImages(7).Picture
   ' Amor...
   .SetItem 7, MensajeRecurso(409), Cliente.ImagenCaras.ListImages(8).Picture
   ' Idea...
   .SetItem 8, MensajeRecurso(410), Cliente.ImagenCaras.ListImages(9).Picture
  End With

 ' **************************************************************
 ' Carga Menu de Tamaños
 ' **************************************************************
 Set MenuDeTamanos = New IcoMenu
  With MenuDeTamanos
   .SetItem 0, "8" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 1, "10" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 2, "12" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 3, "14" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 4, "16" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 5, "18" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 6, "20" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 7, "22" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
   .SetItem 8, "24" & MensajeRecurso(363), Cliente.ImagenTamanos.ListImages(1).Picture
  End With

End Sub
Public Sub CargarMenuEstadosyEliminacion()
Dim Contador, Contador2 As Integer
Dim ItemMenu, Imagen As String

 ' **************************************************************
 ' Carga el Menu con los Estados de los Usuarios
 ' **************************************************************
  Set Me.MenuDeEstadoDeUsuario = New IcoMenu
  Set Me.MenuDeEliminacionDeUsuario = New IcoMenu
  Contador = 0
  For Contador = 1 To CantidadDeAmigosEnChat
   With Me.MenuDeEstadoDeUsuario
    ItemMenu = Trim(AmigosEnChat(Contador).IDAmigoAlias) & " (" & _
               Trim(AmigosEnChat(Contador).Estadotexto) & ")"
    Select Case AmigosEnChat(Contador).EstadoNumerico
     Case "-1"
      Imagen = "UsuarioNoExiste"
     Case "0"
      Imagen = "Desconocido"
     Case "1"
      Imagen = "EstadoVisible"
     Case "2"
      Imagen = "EstadoNoDisponible"
     Case "3"
      Imagen = "EstadoCustom"
    End Select
    .SetItem Contador - 1, ItemMenu, Cliente.Imagenes.ListImages(Imagen).Picture, Trim(AmigosEnChat(Contador).IDAmigoAlias)
    ' Define si hay MultiChat, si lo hay, muestra los Amigos a Eliminar
    ' No se muestra el 1 ya que el mismo corresponde con el cual se
    ' establecio la primera coneccion...
    If CantidadDeAmigosEnChat > 1 Then ' And Contador > 1 Then
     Contador2 = Contador2 + 1
     Me.MenuDeEliminacionDeUsuario.SetItem Contador2 - 1, ItemMenu, Cliente.Imagenes.ListImages(Imagen).Picture, Trim(AmigosEnChat(Contador).IDAmigoAlias)
    End If
   End With
  Next
 ' **************************************************************
  
 ' Si no hay MultiChat, avisa que no hay nada que eliminar
 If CantidadDeAmigosEnChat = 1 Then
  ' No Hay Amigos Para Eliminar...
  Me.MenuDeEliminacionDeUsuario.SetItem 0, MensajeRecurso(411), Cliente.Imagenes.ListImages("UsuarioNoExiste").Picture, ""
 End If

End Sub
Public Sub CargaEstadoIndividual(Usuario As String, Estado As String, Estadotexto As String, Sexo As String)
Dim Contador As Integer
Dim Cambio As Boolean

 ' **************************************************************
 ' Carga los Datos de un Usuario Individual
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  If UCase(Trim(AmigosEnChat(Contador).IDAmigoAlias)) = UCase(Trim(Usuario)) Then
   AmigosEnChat(Contador).Sexo = Sexo
   AmigosEnChat(Contador).EstadoNumerico = Estado
   'Cambio = False
   Select Case Estado
    Case "-1"
     ' Usuario Inexistente...
     'If UCase(Trim(AmigosEnChat(Contador).Estadotexto)) <> MensajeRecurso(342) Then Cambio = True
     AmigosEnChat(Contador).Estadotexto = MensajeRecurso(342)
    Case "0"
     ' No Conectado...
     'If UCase(Trim(AmigosEnChat(Contador).Estadotexto)) <> MensajeRecurso(287) Then Cambio = True
     AmigosEnChat(Contador).Estadotexto = MensajeRecurso(287)
    Case "1"
     ' Disponible...
     'If UCase(Trim(AmigosEnChat(Contador).Estadotexto)) <> MensajeRecurso(180) Then Cambio = True
     AmigosEnChat(Contador).Estadotexto = MensajeRecurso(180)
    Case "2"
     ' No Disponible...
     'If UCase(Trim(AmigosEnChat(Contador).Estadotexto)) <> MensajeRecurso(181) Then Cambio = True
     AmigosEnChat(Contador).Estadotexto = MensajeRecurso(181)
    Case "3"
     'If UCase(Trim(AmigosEnChat(Contador).Estadotexto)) <> Estadotexto Then Cambio = True
     AmigosEnChat(Contador).Estadotexto = Estadotexto
   End Select
   'If Cambio = True Then
     
   'End If
  End If
 Next
      
 MostrarLosEstados
 
End Sub
Private Sub PonerEstadoEnUsuario(NumeroDeAmigoenChat As Integer)
Dim Contador As Integer
Dim Cambio As Boolean

 ' **************************************************************
 ' Recorre todo los Usuario del Listado de Maigos y cuando
 ' Encuentra el NumeroDeAmigoEnChat Completa los Datos
 ' **************************************************************
 Cambio = False
 For Contador = 1 To Variables.CantidadGrupoAmigo
   If UCase(Trim(AmigosEnChat(NumeroDeAmigoenChat).IDAmigoAlias)) = UCase(Trim(Variables.GrupoAmigo(Contador).IDNombreDelAmigo)) Then
    ' Define que se encontro el Amigo dentro del Listado de Amigos...
    Cambio = True
    ' Desconocido...
    AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = MensajeRecurso(412)
    AmigosEnChat(NumeroDeAmigoenChat).Sexo = "D"
    ' Define el Sexo
    AmigosEnChat(NumeroDeAmigoenChat).Sexo = UCase(Trim(Variables.GrupoAmigo(Contador).Sexo))
    Select Case Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado
     Case "0" ' 0. No Conectado
      ' No Conectado...
      AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = MensajeRecurso(287)
      AmigosEnChat(NumeroDeAmigoenChat).EstadoNumerico = "0"
     Case "1" ' 1. Visible Normal
      ' Disponible...
      AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = MensajeRecurso(180)
      AmigosEnChat(NumeroDeAmigoenChat).EstadoNumerico = "1"
     Case "2" ' 2. No Disponible
      ' No Disponible...
      AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = MensajeRecurso(181)
      AmigosEnChat(NumeroDeAmigoenChat).EstadoNumerico = "2"
     Case "3" ' 3. Custom
      AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = Trim(Variables.GrupoAmigo(Contador).EstadoDelAmigoTexto)
      AmigosEnChat(NumeroDeAmigoenChat).EstadoNumerico = "3"
    End Select
    Exit For
   End If
 Next

 
 
 ' **************************************************************
 ' Si no se definio un Estado Entonces Envia un a Query
 ' **************************************************************
 If Cambio = False Then
  ' **************************************************************
  ' No le cambia el Estado, ya que en definitiva lo deja en el estado actual
  ' hasta recibir respuesta del usuario....
  ' **************************************************************
  'AmigosEnChat(NumeroDeAmigoenChat).Estadotexto = MensajeRecurso(412)
  'AmigosEnChat(NumeroDeAmigoenChat).EstadoNumerico = "-1"
  EnviarPaqueteTCP "11" & CompletarCadena(CStr(AmigosEnChat(NumeroDeAmigoenChat).IDAmigoAlias), 16, "D", " ")
 End If

 MostrarLosEstados
 
End Sub

Private Sub Image1_Click()
Dim Posicion As POINTAPI
 
 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + Image1.Left
 Posicion.Y = Me.Top + Image1.HeighT + Image1.Top + 5
 CargarMenuDeMultiChat
 Me.MenuDeEleccionMultiChat.ShowMenu Posicion.X, Posicion.Y

End Sub
Private Sub Label1_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + Label1.Left
 Posicion.Y = Me.Top + Label1.HeighT + Label1.Top + 5
 Me.CargarMenuEstadosyEliminacion
 Me.MenuDeEstadoDeUsuario.ShowMenu Posicion.X, Posicion.Y

End Sub
Private Sub CargarMenuDeMultiChat()
Dim ItemDelMenu, ItemReal, Contador2, Contador As Integer
Dim Bandera As Boolean
Dim ItemMenu, Imagen As String

 ' **************************************************************
 ' Carga el Menu con los Estados de los Usuarios
 ' **************************************************************
  ItemDelMenu = 0
  Set Me.MenuDeEleccionMultiChat = New IcoMenu
  For Contador = 1 To Variables.CantidadGrupoAmigo
   With Me.MenuDeEleccionMultiChat
    ' **************************************************************
    ' Descarta los Que ya existen en el MultiChat
    ' **************************************************************
    Bandera = False
    For Contador2 = 1 To CantidadDeAmigosEnChat
     If UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) = UCase(Trim(AmigosEnChat(Contador2).IDAmigoAlias)) Then
      Bandera = True
      Exit For
     End If
    Next
    ' **************************************************************
    
    ' **************************************************************
    ' Descarta al Porpio Usuario Como Usuario del MultiCHAT
    ' **************************************************************
    If UCase(Trim(Configuracion.IDAliasUsuario)) = UCase(Trim(GrupoAmigo(Contador).IDNombreDelAmigo)) Then
     Bandera = True
    End If
    ' **************************************************************
    
    ' **************************************************************
    ' Solo carga los Amigos que estan Visibles o Custom que
    ' no sea el Mismo Usuario, o que este el la Lista de Multichat...
    ' **************************************************************
    If Not Bandera Then ' Esta Variable es definida por los chequeos de Arriba
     If Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado = 1 Or Variables.GrupoAmigo(Contador).EstadoDelAmigoEstado = 3 Then
      ItemMenu = Trim(GrupoAmigo(Contador).IDNombreDelAmigo) & " ("
      Select Case GrupoAmigo(Contador).EstadoDelAmigoEstado
       Case "1"
        ItemDelMenu = ItemDelMenu + 1
        ' Disponible (Normal)...
        ItemMenu = ItemMenu & MensajeRecurso(180) & ")"
        Imagen = "EstadoVisible"
       Case "3"
        ItemDelMenu = ItemDelMenu + 1
        If Len(GrupoAmigo(Contador).EstadoDelAmigoTexto) > 10 Then
          ItemMenu = ItemMenu & Mid$(GrupoAmigo(Contador).EstadoDelAmigoTexto, 1, 7) & "...)"
         Else
          ItemMenu = ItemMenu & GrupoAmigo(Contador).EstadoDelAmigoTexto & ")"
        End If
        Imagen = "EstadoCustom"
      End Select
      ItemReal = ItemDelMenu - 1
      .SetItem ItemReal, ItemMenu, Cliente.Imagenes.ListImages(Imagen).Picture, Trim(GrupoAmigo(Contador).IDNombreDelAmigo)
     End If
    End If
   End With
  Next
  ' **************************************************************

  ' **************************************************************
  ' No existen Amigos Para Agregar...
  ' **************************************************************
  If ItemDelMenu = 0 Then
   ' No Hay Amigos Disponibles...
   Me.MenuDeEleccionMultiChat.SetItem 0, MensajeRecurso(413), Cliente.Imagenes.ListImages("UsuarioNoExiste").Picture, ""
  End If
  
End Sub

Private Sub Label2_Click()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
  
 ' **************************************************************
 ' Mostrar los Estados
 ' **************************************************************
 Posicion.X = Me.Left + Label2.Left
 Posicion.Y = Me.Top + Label2.HeighT + Label2.Top + 5
 CargarMenuDeMultiChat
 Me.MenuDeEleccionMultiChat.ShowMenu Posicion.X, Posicion.Y

End Sub
Sub EstadoAnimacion(Estado As Boolean)

 ' **************************************************************
 ' Define el Estado de la Animacion cuando se estan transmitiendo
 ' datos...
 ' **************************************************************
 If Estado Then
   Animacion.Enabled = True
  Else
   Animacion.Enabled = False
   AnimacionImagen.Picture = Cliente.AnimacionDeConeccion.ListImages(1).Picture
 End If
End Sub
Sub CargarMenuEleccionDeUsuario(Usuario As String)
 
 ' **************************************************************
 ' Carga el Menu con los Estados de los Usuarios
 ' **************************************************************
  Set Me.MenuBLoqueo = New IcoMenu
  With Me.MenuBLoqueo
   ' Bloquear Amigo...
   .SetItem 0, MensajeRecurso(304), Cliente.Imagenes.ListImages("UsuarioNoExiste").Picture, Trim(Usuario)
   ' Mensaje Privado...
   .SetItem 1, MensajeRecurso(415), Cliente.Imagenes.ListImages("Mensaje").Picture, Trim(Usuario)
   ' Desbloquear Amigo...
   .SetItem 2, MensajeRecurso(305), Cliente.Imagenes.ListImages("EstadoVisible").Picture, Trim(Usuario)
  End With
  ' **************************************************************

  ' **************************************************************
  ' No existen Amigos Para Agregar...
  ' **************************************************************
  'If ItemDelMenu = 0 Then
  ' Me.MenuDeEleccionMultiChat.SetItem 0, "No Hay Amigos Disponibles...", Cliente.Imagenes.ListImages("UsuarioNoExiste").Picture, ""
  'End If
  
End Sub
Function UsuarioBloqueado(Nombre As String) As Boolean
Dim Contador As Integer

  ' **************************************************************
  ' Verifica si el Usuario esta Bloqueado
  ' **************************************************************
  If Variables.UsuarioBloqueadoCantidad = 0 Then
   UsuarioBloqueado = False
   Exit Function
  End If
  
  ' **************************************************************
  ' Verifica si el Usuario esta Bloqueado
  ' **************************************************************
  For Contador = 1 To Variables.UsuarioBloqueadoCantidad
   If UCase(Trim(Nombre)) = UCase(Trim(Variables.UsuarioBloqueadoNombres(Contador).NombreDelUsuario)) Then
    ' Usuario Bloqueado
    UsuarioBloqueado = True
    Exit Function
   End If
  Next
 
  ' **************************************************************
  ' Define que no esta Bloqueado
  ' **************************************************************
  UsuarioBloqueado = False
  
End Function

Private Sub Picture2_Click()

 Scroller.ScrollListBox Me.ListadoAmigosMultiChat, "Arriba"
 
End Sub

Private Sub Picture3_Click()

 Scroller.ScrollListBox Me.ListadoAmigosMultiChat, "Abajo"

End Sub

Private Sub ScrollAbajo_Click()
Dim Respuesta As Variant
 
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
    
 
 Respuesta = ScrollText&(VentanaMensajes, 1)

End Sub

Private Sub ScrollArriba_Click()
Dim Respuesta As Variant
    
 ' **************************************************************
 ' Ejecutar Sonido
 ' **************************************************************
 EjecutarSonido "003"
    
 Respuesta = ScrollText&(VentanaMensajes, -1)
    
End Sub
Private Sub Timer1_Timer()

 FlashWindow Me.hwnd, 1
 Me.Timer1.Enabled = False

End Sub
Public Function GetHyperlink(rch As RichTextBox, X As Single, Y As Single) As String
    Dim pt As POINTAPI__
    Dim pos As Integer
    Dim ch As String
    Dim txt As String
    Dim txtlen As Integer
    Dim pos_start As Integer
    Dim pos_mijloc As Integer
    Dim pos_end As Integer
    
    ' convert mouse pos in pixels
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' position of character under cursor
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then
        Exit Function
    End If
    txt = rch.Text

    ' get start position of word under cursor
    For pos_start = pos To 1 Step -1
        If Mid$(txt, pos_start + 1, 1) = Chr(13) Then
            rch.ToolTipText = ""
            rch.MousePointer = 0
            Exit Function
        End If
        ch = Mid$(txt, pos_start, 1)
        If ch = Chr(32) Or ch = vbCr Or ch = vbLf Or ch = vbNewLine Then Exit For
    Next pos_start
    pos_start = pos_start + 1

    ' get end position of word under cursor
    txtlen = Len(txt)
    For pos_end = pos To txtlen
        ch = Mid$(txt, pos_end, 1)
    If ch = Chr(32) Or ch = vbCr Then Exit For
    Next pos_end
    pos_end = pos_end - 1

    If pos_start <= pos_end Then _
        GetHyperlink = Mid$(txt, pos_start, pos_end - pos_start + 1)
        
        
        
        If Left(GetHyperlink, 5) = "http:" Or Left(GetHyperlink, 4) = "www." Or Left(GetHyperlink, 7) = "mailto:" Then
            rch.MousePointer = vbCustom
            If Left(GetHyperlink, 7) <> "mailto:" Then
                        ' Haga Click Aqui para Navegar a
                        rch.ToolTipText = MensajeRecurso(417) + GetHyperlink
            Else
                        ' Haga Click Aqui para Enviar un Mail a
                        rch.ToolTipText = MensajeRecurso(418) + Right(GetHyperlink, Len(GetHyperlink) - 7)
            End If
        ElseIf InStr(1, GetHyperlink, "@") > 0 Then
            'rchHyperlink.MouseIcon = LoadResPicture(101, vbResCursor)
            rch.MousePointer = vbCustom
            rch.ToolTipText = MensajeRecurso(418) + GetHyperlink
        Else
            rch.ToolTipText = ""
            rch.MousePointer = 0
        End If
               
        
End Function
Private Sub VentanaMensajes_Click()
Dim lngRet As Long
Dim Htxt As String

   ' **************************************************************
   ' Habre el Browser y/o el Cliente Mail segun corresponda...
   ' **************************************************************
   Htxt = HiperVinculo
   If Left(Htxt, 5) = "http:" Or Left(Htxt, 7) = "mailto:" Or Left(Htxt, 4) = "www." Then
    lngRet = ShellExecute(0&, "Open", Htxt, "", vbNullString, 1)
   End If
   If InStr(1, Htxt, "@") > 0 Then lngRet = ShellExecute(0&, "Open", "mailto:" + Htxt, "", vbNullString, 1)

End Sub
Private Sub VentanaMensajes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HiperVinculo = GetHyperlink(VentanaMensajes, X, Y)
    
End Sub
Public Sub AvisarCambioDeEstado(Usuario As String) ', EstadoNumero As Integer, Estadotexto As String)
Dim Respuesta As Integer
    
  ' **************************************************************
  ' Verifica que el Amigo este...
  ' **************************************************************
  Respuesta = TratarAmigosEnChat("Buscar", Trim(Usuario))
  If Respuesta = 0 Then Exit Sub ' El Amigo no esta en esta ventana de Mensajes...
  
  Me.AgregarLineaGrisConDatos (Respuesta)
 
End Sub
Public Function AgregarEventoDelUsuario(Evento As String)

 ' **************************************************************
 ' Procesa los eventos tipo... Esta escribiendo una respuesta..
 ' etc.
 ' **************************************************************
 If Trim(Evento) = "" Then
   Me.LabelMensajeEvento.ForeColor = &H8000000F
   Me.EventosDeUsuarioTexto = MensajeRecurso(351)
   Me.LabelMensajeEvento = Me.EventosDeUsuarioTexto
  Else
   Me.EventosDeUsuarioTexto = Evento
   ' Restartea el contador de Eventos...
   Me.EventosDeusuarios.Enabled = False
   Me.EventosDeusuarios.Enabled = True
   ' Muestra la Imagen y el Texto
   Me.LabelMensajeEvento.ForeColor = vbBlue
   Me.LabelMensajeEvento = Me.EventosDeUsuarioTexto
 End If
  
End Function
Public Function FunctionAlguienDejoElMultichat(Usuario As String)
 
 ' **************************************************************
 ' EVENTO 2: Si se recibe este evento, envia el Listado de
 ' participantes de Multichat...
 ' **************************************************************
 If CantidadDeAmigosEnChat > 1 And ResponsableMultiChat = True Then
  ' Saca el Usuario que se desconecto...
  Me.TratarAmigosEnChat "Sacar", Trim(Usuario)
  ' Como es el Owner envia el listado Multichat....
  Me.EnviarListadoDeMultiChat
 End If

End Function
Public Function CancelarMultichat(Mensaje As String)
 
 ' **************************************************************
 ' EVENTO 3: Si se recibe este evento, se debe cerrar el Multichat
 ' ya que el Owner lo cancelo...
 ' **************************************************************
 CantidadDeAmigosEnChat = 1 ' Con este cambio pasa a Chat Simple...
 ' **************************************************************
 ' Pone el Titulo de la Ventana
 ' **************************************************************
 PonerLabels
 ' **************************************************************
 ' Pone el Estado del Usuario o Multichat...
 ' **************************************************************
 MostrarLosEstados
 ' **************************************************************
 ' Carga el Listado de MultiChat
 ' **************************************************************
 CargarAmigosMultiChat
 ' **************************************************************
 ' Igualmente lo avisa con un MSGBOX
 MostrarMSGBox Mensaje, vbOKOnly, "vbCritical", Configuracion.TituloVentanas
 
End Function
Public Sub EnviarListadoDeMultiChat(Optional Nombre As String, Optional VentanaID As String)
Dim Contador, Cantidad As Integer
Dim AmigosMultiChat  As String
Dim PaqueteEnviar As String
Dim ComandoAdicional As String

 ' **************************************************************
 ' Dispara la Animacion
 ' **************************************************************
  EstadoAnimacion True
 
 ' **************************************************************
 ' Enviar el Listado a Todos los usuarios...
 ' **************************************************************
 For Contador = 1 To CantidadDeAmigosEnChat
  ComandoAdicional = ComandoAdicional & CompletarCadena(AmigosEnChat(Contador).IDAmigoAlias, 16, "D", " ") & CompletarCadena(AmigosEnChat(Contador).IDVentanaAmigo, 10, "I", "0")
 Next
 ' Se Agrega a si Mismo...
 ComandoAdicional = ComandoAdicional & CompletarCadena(Trim(Configuracion.IDAliasUsuario), 16, "D", " ") & CompletarCadena(Me.hwnd, 10, "I", "0")
  
 ' **************************************************************
 ' Para cuando se Borra un Usuario Multichat
 ' **************************************************************
 If Trim(Nombre) <> "" Then
  PaqueteEnviar = "3" & CompletarCadena(Trim(Nombre), 16, "D", " ") & "4" & VentanaID & CompletarCadena(CStr(CantidadDeAmigosEnChat + 1), 2, "I", "0") & ComandoAdicional
  ' **************************************************************
  ' Envia el Paquete
  ' **************************************************************
  EnviarPaqueteTCP PaqueteEnviar
  EstadoAnimacion False
  Exit Sub
 End If
 
 ' **************************************************************
 ' Envia el Paquete a TODOS los Amigos...
 ' **************************************************************
 ComandoAdicional = "M" & CompletarCadena(CStr(CantidadDeAmigosEnChat + 1), 2, "I", "0") & ComandoAdicional
 EnviarPaqueteTCP "44" & CompletarCadena(Configuracion.IDAliasUsuario, 16, "D", " ") & CompletarCadena(Me.hwnd, 10, "I", "0") & _
                   ComandoAdicional
 
 ' **************************************************************
 ' Frena la Animacion
 ' **************************************************************
 EstadoAnimacion False

End Sub


