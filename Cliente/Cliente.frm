VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Cliente 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   5805
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4020
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Cliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "Cliente.frx":058A
   Picture         =   "Cliente.frx":0894
   ScaleHeight     =   5805
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImagenesSonido 
      Left            =   8220
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":3899
            Key             =   "ConSonido"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":3C33
            Key             =   "SinSonido"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ImagenAviso 
      Height          =   255
      Left            =   4800
      Picture         =   "Cliente.frx":3FCD
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   47
      Top             =   3420
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox LineaGris 
      Height          =   225
      Left            =   5340
      Picture         =   "Cliente.frx":43A9
      ScaleHeight     =   165
      ScaleWidth      =   1875
      TabIndex        =   46
      Top             =   2940
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImagenesFlechaFinas 
      Left            =   6930
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":4446
            Key             =   "FinaIzquierda"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":47BE
            Key             =   "FinaAbajo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":4B4A
            Key             =   "FinaArriba"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":4ED7
            Key             =   "FinaDerecha"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox IconoFormularioXXX 
      Height          =   375
      Left            =   4830
      Picture         =   "Cliente.frx":524F
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   2940
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash6 
      Height          =   315
      Left            =   7020
      Picture         =   "Cliente.frx":57D9
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   44
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash6 
      Height          =   315
      Left            =   7020
      Picture         =   "Cliente.frx":5D63
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   43
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash6 
      Height          =   315
      Left            =   7020
      Picture         =   "Cliente.frx":62ED
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   42
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash5 
      Height          =   315
      Left            =   6570
      Picture         =   "Cliente.frx":6877
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   41
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash5 
      Height          =   315
      Left            =   6570
      Picture         =   "Cliente.frx":6E01
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   40
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash5 
      Height          =   315
      Left            =   6570
      Picture         =   "Cliente.frx":738B
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash4 
      Height          =   315
      Left            =   6120
      Picture         =   "Cliente.frx":7915
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   38
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash4 
      Height          =   315
      Left            =   6120
      Picture         =   "Cliente.frx":7E9F
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   37
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash4 
      Height          =   315
      Left            =   6120
      Picture         =   "Cliente.frx":8429
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImagenesFlecha 
      Left            =   7560
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":89B3
            Key             =   "ArribaAzul"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":8D40
            Key             =   "AbajoAzul"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":90CC
            Key             =   "AbajoRoja"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":91BF
            Key             =   "AbajoVerde"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox IconoTrayConectadoSinMensajeRojo 
      Height          =   315
      Left            =   8730
      Picture         =   "Cliente.frx":92B1
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   35
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoSinMensajeAmarillo 
      Height          =   315
      Left            =   8310
      Picture         =   "Cliente.frx":983B
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   34
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoSinMensajeVerde 
      Height          =   315
      Left            =   7890
      Picture         =   "Cliente.frx":9DC5
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   33
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash3 
      Height          =   315
      Left            =   5670
      Picture         =   "Cliente.frx":A34F
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash3 
      Height          =   315
      Left            =   5670
      Picture         =   "Cliente.frx":A8D9
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   31
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash3 
      Height          =   315
      Left            =   5670
      Picture         =   "Cliente.frx":AE63
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   30
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash2 
      Height          =   315
      Left            =   5220
      Picture         =   "Cliente.frx":B3ED
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash2 
      Height          =   315
      Left            =   5220
      Picture         =   "Cliente.frx":B977
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash2 
      Height          =   315
      Left            =   5220
      Picture         =   "Cliente.frx":BF01
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImagenesMenus 
      Left            =   6300
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":C48B
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":CA25
            Key             =   "Mensaje"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":CDBF
            Key             =   "EnviarMail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":D2E1
            Key             =   "CambiarEstado"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":D87B
            Key             =   "EnviarArchivo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":DE15
            Key             =   "MoverAGrupo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":E3AF
            Key             =   "MensajeGrupal"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":E949
            Key             =   "Preferencias"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":EEE3
            Key             =   "CambiarMisDatos"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":F27D
            Key             =   "RecargarAmigos"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":F817
            Key             =   "MostrarMisDatos"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":FBB1
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":FF4B
            Key             =   "EIM"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":104E5
            Key             =   "Clave"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":10A7F
            Key             =   "GrupoEliminar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":11019
            Key             =   "RenombrarGrupo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":115B3
            Key             =   "GrupoAgregar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":11B4D
            Key             =   "BuscarAmigos"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":120E7
            Key             =   "BloqueoDeAmigos"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":12681
            Key             =   "AmigoAgregar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":12C1B
            Key             =   "AmigoEliminar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":131B5
            Key             =   "CambiarUsuario"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1374F
            Key             =   "Desconectar"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":13AE9
            Key             =   "Conectar"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox IconoAplicacion 
      Height          =   375
      Left            =   4320
      Picture         =   "Cliente.frx":13E83
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   2940
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   180
      Picture         =   "Cliente.frx":1440D
      ScaleHeight     =   255
      ScaleWidth      =   3435
      TabIndex        =   24
      Top             =   4360
      Width           =   3435
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   3380
      ScaleHeight     =   2715
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   1830
      Width           =   235
   End
   Begin VB.PictureBox ScrollDerecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   145
      Left            =   3500
      MouseIcon       =   "Cliente.frx":14CBC
      MousePointer    =   99  'Custom
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   23
      Top             =   4685
      Width           =   145
   End
   Begin VB.PictureBox ScrollIzquierda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   145
      Left            =   150
      MouseIcon       =   "Cliente.frx":14E0E
      MousePointer    =   99  'Custom
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   22
      Top             =   4685
      Width           =   145
   End
   Begin VB.PictureBox ScrollAbajo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   145
      Left            =   3720
      MouseIcon       =   "Cliente.frx":14F60
      MousePointer    =   99  'Custom
      ScaleHeight     =   172.222
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   21
      Top             =   4680
      Width           =   145
   End
   Begin VB.PictureBox ScrollArriba 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   145
      Left            =   3720
      MouseIcon       =   "Cliente.frx":150B2
      MousePointer    =   99  'Custom
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   20
      Top             =   1820
      Width           =   145
   End
   Begin VB.Timer AvisoMensajesPendientes 
      Interval        =   65000
      Left            =   6300
      Top             =   870
   End
   Begin VB.PictureBox IconoTrayConectandoConMensajeFlash 
      Height          =   315
      Left            =   4830
      Picture         =   "Cliente.frx":15204
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectandoConMensaje 
      Height          =   315
      Left            =   4320
      Picture         =   "Cliente.frx":1578E
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensajeFlash 
      Height          =   315
      Left            =   4800
      Picture         =   "Cliente.frx":15D18
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectadoConMensaje 
      Height          =   315
      Left            =   4320
      Picture         =   "Cliente.frx":162A2
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoConMensajeFlash 
      Height          =   315
      Left            =   4800
      Picture         =   "Cliente.frx":1682C
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayDesConectado 
      Height          =   315
      Left            =   7470
      Picture         =   "Cliente.frx":16DB6
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectando 
      Height          =   315
      Left            =   7470
      Picture         =   "Cliente.frx":17340
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   2430
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox IconoTrayConectadoSinMensaje 
      Height          =   315
      Left            =   7470
      Picture         =   "Cliente.frx":178CA
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer TimerMensaje 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   870
   End
   Begin VB.Timer RefrescoAmigos 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   4710
      Top             =   870
   End
   Begin MSComctlLib.TreeView ListadoDeAmigos 
      Height          =   2775
      Left            =   180
      TabIndex        =   4
      Top             =   1830
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   4895
      _Version        =   393217
      Indentation     =   317
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   3
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImagenesAmigos"
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Cliente.frx":17E54
   End
   Begin MSWinsockLib.Winsock TCPSocket 
      Left            =   4740
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer OnLineTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5190
      Top             =   870
   End
   Begin VB.PictureBox IconoTrayConectadoConMensaje 
      Height          =   315
      Left            =   4320
      Picture         =   "Cliente.frx":17FB6
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList AnimacionDeConeccion 
      Left            =   6960
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":18540
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":185FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":186C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":18787
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImagenTamanos 
      Left            =   7590
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":18849
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImagenCaras 
      Left            =   8190
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":18BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":18FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":19439
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":19865
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":19CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1A103
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1A530
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1A95B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1AD74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ListadoImagenes 
      Left            =   6960
      Top             =   3750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1B190
            Key             =   "vbInformation"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1BE6A
            Key             =   "vbQuestion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1CB44
            Key             =   "vbWarning"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1D81E
            Key             =   "vbCritical"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImagenesAmigos 
      Left            =   6270
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1E4F8
            Key             =   "Grupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1EA92
            Key             =   "M0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1F02C
            Key             =   "M1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1F5C6
            Key             =   "M2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":1FB60
            Key             =   "M3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":200FA
            Key             =   "F0"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":20694
            Key             =   "F1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":20C2E
            Key             =   "F2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":211C8
            Key             =   "F3"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":21762
            Key             =   "UsuarioNoExiste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":21CFC
            Key             =   "CargandoAmigos"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":22296
            Key             =   "Hombre"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":22830
            Key             =   "Mujer"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":22DCA
            Key             =   "Conectando"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":23364
            Key             =   "Desconectado"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":238FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":23E98
            Key             =   "Mensaje0"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":24432
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":249CC
            Key             =   "Mensaje"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":24F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":25500
            Key             =   "MensajeFlash"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":25A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":26034
            Key             =   "MensajeFlash2"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":265CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":26B68
            Key             =   "MensajeFlash3"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":27102
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2769C
            Key             =   "MensajeFlash4"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":27C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":281D0
            Key             =   "MensajeFlash5"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2876A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":28D04
            Key             =   "MensajeFlash6"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2929E
            Key             =   "ListadoUsuario"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":29838
            Key             =   "Desconocido"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":29DD2
            Key             =   "UsuarioInexistente"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2A36C
            Key             =   "MultiChat"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   6300
      Top             =   3750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2A906
            Key             =   "Conectando"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2ACA0
            Key             =   "Desconectado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2B03A
            Key             =   "Conectado"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2B3D4
            Key             =   "EstadoNoDisponible"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2B96E
            Key             =   "EstadoCustom"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2BF08
            Key             =   "EstadoVisible"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2C4A2
            Key             =   "NoConectado"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2CA3C
            Key             =   "Desconocido"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2CFD6
            Key             =   "UsuarioNoExiste"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2D570
            Key             =   "MultiChat"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2DB0A
            Key             =   "Mensaje"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2E0A4
            Key             =   "Mujer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2E63E
            Key             =   "Hombre"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2EBD8
            Key             =   "UsuarioBloqueado"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2F172
            Key             =   "Minimizar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2F3CD
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2F62B
            Key             =   "Refrescando"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":2FBC5
            Key             =   "Grupo"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3350
      MouseIcon       =   "Cliente.frx":3015F
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3690
      MouseIcon       =   "Cliente.frx":302B1
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   240
   End
   Begin VB.Label DropArea 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   3270
   End
   Begin VB.Label TituloVentana1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   120
      Width           =   2115
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   90
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "4020"
      Height          =   195
      Left            =   4740
      TabIndex        =   10
      Top             =   5250
      Width           =   1125
   End
   Begin VB.Image SonidoSeteo 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3640
      MouseIcon       =   "Cliente.frx":30403
      MousePointer    =   99  'Custom
      Top             =   5020
      Width           =   270
   End
   Begin VB.Label MenuBotonAyuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3000
      MouseIcon       =   "Cliente.frx":30555
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   480
      Width           =   585
   End
   Begin VB.Label MenuBotonAmigos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2310
      MouseIcon       =   "Cliente.frx":306A7
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   480
      Width           =   645
   End
   Begin VB.Label MenuBotonConfiguracion 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1140
      MouseIcon       =   "Cliente.frx":307F9
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label MenuBotonConectar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   180
      MouseIcon       =   "Cliente.frx":3094B
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   480
      Width           =   915
   End
   Begin VB.Image ContextualAgregarAyuda 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   3165
      MouseIcon       =   "Cliente.frx":30A9D
      MousePointer    =   99  'Custom
      Picture         =   "Cliente.frx":30BEF
      Top             =   930
      Width           =   660
   End
   Begin VB.Image ContextualAgregarAmigo 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   2430
      MouseIcon       =   "Cliente.frx":31929
      MousePointer    =   99  'Custom
      Picture         =   "Cliente.frx":31A7B
      Top             =   950
      Width           =   630
   End
   Begin VB.Image ContextualAgregarGrupo 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   1695
      MouseIcon       =   "Cliente.frx":3274D
      MousePointer    =   99  'Custom
      Picture         =   "Cliente.frx":3289F
      Top             =   950
      Width           =   630
   End
   Begin VB.Image ContextualRefrescarAmigos 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   945
      MouseIcon       =   "Cliente.frx":33571
      MousePointer    =   99  'Custom
      Picture         =   "Cliente.frx":336C3
      Top             =   950
      Width           =   630
   End
   Begin VB.Image ImageMenuDeEstado 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   205
      MouseIcon       =   "Cliente.frx":34395
      MousePointer    =   99  'Custom
      Picture         =   "Cliente.frx":344E7
      Top             =   950
      Width           =   630
   End
   Begin VB.Label EstadoUsuarioTexto 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   495
      TabIndex        =   2
      Top             =   5070
      Width           =   2865
   End
   Begin VB.Image EstadoUsuarioImagen 
      Height          =   240
      Left            =   145
      MouseIcon       =   "Cliente.frx":351B9
      MousePointer    =   99  'Custom
      Top             =   5030
      Width           =   240
   End
   Begin VB.Label TiempoEnLinea 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3090
      MouseIcon       =   "Cliente.frx":3530B
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5460
      Width           =   795
   End
   Begin VB.Image EstadoCLienteImagen 
      Height          =   165
      Left            =   150
      Top             =   5440
      Width           =   165
   End
   Begin VB.Label EstadoClienteTexto 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   0
      Top             =   5460
      Width           =   2295
   End
End
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
' Variables Varias...
' **************************************************************
Public UltimoClickDeAmigo As String
Public IndiceTimerMensaje As Integer
Public MenuClickHandle As Long
Const EspacioArriba = 700
Const EspacioIzquierda = 0

' **************************************************************
' Tray Icon...
' **************************************************************
Public WithEvents SysIcon As CSystrayIcon  'Create an instance of CSystrayIcon using events
Attribute SysIcon.VB_VarHelpID = -1

' **************************************************************
' Menus (Todos)...
' **************************************************************
Public WithEvents MenuCambioDeEstado As IcoMenu
Attribute MenuCambioDeEstado.VB_VarHelpID = -1
Public WithEvents MenuDeCambioDeGrupo As IcoMenu
Attribute MenuDeCambioDeGrupo.VB_VarHelpID = -1
Public WithEvents MenuDeEstadosDeUsuario As IcoMenu
Attribute MenuDeEstadosDeUsuario.VB_VarHelpID = -1
Public WithEvents MenuToolConeccion As IcoMenu
Attribute MenuToolConeccion.VB_VarHelpID = -1
Public WithEvents MenuToolConfiguracion As IcoMenu
Attribute MenuToolConfiguracion.VB_VarHelpID = -1
Public WithEvents MenuToolAmigos As IcoMenu
Attribute MenuToolAmigos.VB_VarHelpID = -1
Public WithEvents MenuToolAyuda As IcoMenu
Attribute MenuToolAyuda.VB_VarHelpID = -1
Public WithEvents MenuClickAmigo As IcoMenu
Attribute MenuClickAmigo.VB_VarHelpID = -1
Public WithEvents MenuClickGrupo As IcoMenu
Attribute MenuClickGrupo.VB_VarHelpID = -1
Public WithEvents MenuClickUsuario As IcoMenu
Attribute MenuClickUsuario.VB_VarHelpID = -1
Public WithEvents MenuClickUsuarioDesconectadoConectando As IcoMenu
Attribute MenuClickUsuarioDesconectadoConectando.VB_VarHelpID = -1
Public WithEvents MenuPopUpTray As IcoMenu
Attribute MenuPopUpTray.VB_VarHelpID = -1
Public WithEvents MenuMensajesPendientes As IcoMenu
Attribute MenuMensajesPendientes.VB_VarHelpID = -1
Private Sub AvisoMensajesPendientes_Timer()
Dim Respuesta, Diferencia As Long
Dim Contador As Integer

 ' **************************************************************
 ' Avisa que no esta conectado...
 ' **************************************************************
 If Configuracion.Logueado <> 3 And Variables.FormularioLoguin = False Then
  Respuesta = MostrarMSGBox(MensajeRecurso(447), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
  If Respuesta = vbYes Then
   ' Llama al metodo estandar para la coneccion..
   MenuToolConeccion_Click 0, ""
  End If
 End If
 
 ' **************************************************************
 ' Procesa el Aviso de Mensajes Pendientes
 ' **************************************************************
 If AvisoMensajesPendientes Then
  ' Si no hay mensajes pendientes sale...
  If CantidadDeMensajesPendientes = 0 Then Exit Sub
  ' Si el Aviso ya fue generado tambien sale...
  If Variables.UltimoMensajePendienteAviso = True Then Exit Sub
  ' Sino, si paso mas de 1 minuto muestra el Aviso
  'Diferencia = DateDiff("s", Variables.UltimoMensajePendienteHorario, Time)
  'If CLng(Diferencia) >= CLng(60) Then
  If CLng(DateDiff("s", Variables.UltimoMensajePendienteHorario, Time)) >= CLng(60) Then
   ' Muestra el Aviso
   Variables.UltimoMensajePendienteAviso = True
   ' Define si es un mensaje o varios (Es importante por el texto en si, la
   ' s en mensaje(s) pendiente(s)
   If CantidadDeMensajesPendientes = 1 Then
     ' Muestra: Usted tiene [ % ] Mensaje Pendiente...¿Desea Ver el Mensaje?
     Respuesta = MostrarMSGBox(MensajeRecurso(170) & CantidadDeMensajesPendientes & MensajeRecurso(171), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
    Else
     ' Muestra: Usted tiene [ % ] Mensajes Pendientes...¿Desa Ver los Mensajes?
     Respuesta = MostrarMSGBox(MensajeRecurso(170) & CantidadDeMensajesPendientes & MensajeRecurso(172), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
   End If
   Variables.UltimoMensajePendienteAviso = False
   If Respuesta = vbNo Then Exit Sub
   ' Carga el Menu de Pendientes
   CargarMenuMensajesPendientes
   ' Muestra los Mensajes
   For Contador = 1 To CantidadDeMensajesPendientesAgrupados
    MenuMensajesPendientes_Click Contador - 1, ""
   Next
  End If
 End If
 
End Sub
Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Permite hacer el Move del Formulario...
 ' **************************************************************
 If Button = 1 Then
  ReleaseCapture
  'Debug.Print "Si"
  'Me.TimerDock.Enabled = True
  SendMessage Me.hwnd, &HA1, 2, 0
  Inicializar.PosicionCliente "Grabar", Me.Left, Me.Top
  'Me.TimerDock.Enabled = False
  'Debug.Print "No"
  Exit Sub
 End If

End Sub
Public Sub CargarTextos()

 ' **************************************************************
 ' Define los Textos del Formulario
 ' **************************************************************
 ' ToolTips...
 Me.ContextualRefrescarAmigos.ToolTipText = MensajeRecurso(166) ' Refresca Su Listado de Amigos...
 Me.ContextualAgregarGrupo.ToolTipText = MensajeRecurso(167) ' Crear Grupo...
 Me.ContextualAgregarAmigo.ToolTipText = MensajeRecurso(130) ' Agregar Amigo...
 Me.ContextualAgregarAyuda.ToolTipText = MensajeRecurso(169) ' Buscar Amigos...
 Me.ImageMenuDeEstado.ToolTipText = MensajeRecurso(165) ' Cambiar Su Estado Actual...
 ' Titulo Ventana...
 Me.TituloVentana1.ForeColor = Variables.FontTituloVentana
 Me.TituloVentana1 = Trim(Configuracion.TituloVentanas) & "..."
 ' Menus Descolgables...
 Me.MenuBotonConectar.ForeColor = Variables.FontMenuDescolgable
 Me.MenuBotonConectar.BackColor = Variables.FontFOndoMenuDescolgable
 Me.MenuBotonConectar = " " & MensajeRecurso(161) ' Conección
 Me.MenuBotonConfiguracion.ForeColor = Variables.FontMenuDescolgable
 Me.MenuBotonConfiguracion.BackColor = Variables.FontFOndoMenuDescolgable
 Me.MenuBotonConfiguracion = " " & MensajeRecurso(162) ' Configuración
 Me.MenuBotonAmigos.ForeColor = Variables.FontMenuDescolgable
 Me.MenuBotonAmigos.BackColor = Variables.FontFOndoMenuDescolgable
 Me.MenuBotonAmigos = " " & MensajeRecurso(163) ' Amigos
 Me.MenuBotonAyuda.ForeColor = Variables.FontMenuDescolgable
 Me.MenuBotonAyuda.BackColor = Variables.FontFOndoMenuDescolgable
 Me.MenuBotonAyuda = " " & MensajeRecurso(164) ' Ayuda
 ' Imagenes...
 Me.Image5.Picture = Cliente.IconoAplicacion.Picture
 Me.Image2.Picture = Cliente.Imagenes.ListImages("Minimizar").Picture
 Me.Image3.Picture = Cliente.Imagenes.ListImages("Cerrar").Picture
 Me.ScrollDerecha.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaDerecha").Picture
 Me.ScrollIzquierda.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaIzquierda").Picture
 Me.ScrollArriba.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaArriba").Picture
 Me.ScrollAbajo.Picture = Cliente.ImagenesFlechaFinas.ListImages("FinaAbajo").Picture
 ' Otros labels
 Me.EstadoUsuarioTexto.ForeColor = Variables.FontFormClienteTextosDeEstado
 Me.EstadoClienteTexto.ForeColor = Variables.FontFormClienteTextosDeEstado
 Me.TiempoEnLinea.ForeColor = Variables.FontFormClienteTimer
 
End Sub
Private Sub EstadoUsuarioImagen_Click()

 ' **************************************************************
 ' Sonido Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Si no esta logueado loinforma y sale...
 ' **************************************************************
 If Configuracion.Logueado <> 3 Then
  ' Muestra: No se Encuentra Logueado al Sistema...
  MostrarMSGBox MensajeRecurso(197), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Muestra el Menu Descolgable...
 ' **************************************************************
 MenuDeEstadosDeUsuario.ShowMenu Me.EstadoUsuarioImagen.Left + Cliente.Left - 20, Me.EstadoUsuarioImagen.Top + Cliente.Top + Me.EstadoUsuarioImagen.HeighT

End Sub
Private Sub Image2_Click() ' Boton Minimizar...

 ' **************************************************************
 ' Ejecuta el Sonido del Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Al estar activo el Icon Tray con el Hide, queda Minimizado...
 ' **************************************************************
 Me.Hide
 
End Sub
Private Sub Image3_Click()

 ' **************************************************************
 ' Ejecuta el Sonido del Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Salir !!!
 ' **************************************************************
 Varios.SalirDelSistema
 
End Sub
Private Sub ListadoDeAmigos_Collapse(ByVal Node As MSComctlLib.Node)

 ' **************************************************************
 ' Si se pica en un Nodo entonces deja que se expanda o calapse
 ' **************************************************************
 ' Solo los nodos que son de Grupo...
 If UCase(Mid$(Node.key, 1, 1)) = "G" And UCase(Trim(UCase(Node.key))) <> "USUARIO" Then
  Varios.CargarEstadoDeNodos (Mid$(Node.key, 2))
  Exit Sub ' Sale sin dejar colapsar...
 End If
 
 ' **************************************************************
 ' No deja Colapsar el Nodo
 ' **************************************************************
 Node.Expanded = True
 
End Sub
Private Sub ListadoDeAmigos_DblClick()
  Dim Respuesta, Contador As Integer
  Dim NodoKey As String
  On Error GoTo SalirError
  
  ' **************************************************************
  ' Ejecuta el Sonido
  ' **************************************************************
  Audio.EjecutarSonido "003"
  
  ' **************************************************************
  ' Cerrar los menus que pueden llegar a abrirse por pensar que
  ' es un Click !...
  ' **************************************************************
  Varios.CerrarMenusDecolgables
   
  ' **************************************************************
  ' Carga el Nodo Key
  ' **************************************************************
  NodoKey = Me.ListadoDeAmigos.SelectedItem.key
  
  ' **************************************************************
  ' Dispara los Mensajes pendientes
  ' **************************************************************
  If Mid$(Me.ListadoDeAmigos.SelectedItem.Text, 1, 1) = "*" Then
   Select Case Variables.CantidadDeMensajesPendientes
    Case 0
     ' **************************************************************
     ' No posee mensaje pendientes
     ' **************************************************************
     ' Muestra: Usted no Posee Mensajes Pendientes...
     MostrarMSGBox MensajeRecurso(194), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    Case Else
     ' Carga el Menu de Pendientes
     CargarMenuMensajesPendientes
     If CantidadDeMensajesPendientesAgrupados = 1 Then
       ' Si hay un solo Mensaje lo habre...
       MenuMensajesPendientes_Click 0, ""
      Else
       ' Si no avisa...
       ' Muestra: Posee Mensajes de Mas de 1 Usuario. ¿Desea Ver todo los Mensajes?
       Respuesta = MostrarMSGBox(MensajeRecurso(173), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
       If Respuesta = vbYes Then
        ' Muestra todos los Mensajes
        For Contador = 1 To CantidadDeMensajesPendientesAgrupados
         MenuMensajesPendientes_Click Contador - 1, ""
        Next
       End If
     End If
   End Select
   Exit Sub ' Sale ya que acontinuacion se trata el nodo de Amigo...
  End If
   
  ' **************************************************************
  ' Verifica que sea un Usuario
  ' **************************************************************
  If UCase(Mid$(NodoKey, 1, 1)) = "U" And UCase(Mid$(NodoKey, 1, 7)) <> "USUARIO" Then
    ' **************************************************************
    ' Si esta conectado, Habre la Ventana correspondiente
    ' **************************************************************
    ' Definicion de Caratceres en NodeKey
    ' 1ro. Caracter
    '      Tipo de Nodo (U=usuario, G=Grupo,*=MensajesPendientes)
    ' 2do. caracter
    '      0. No Conectado
    '      1. Visible Normal
    '      2. No Disponible
    '      3. Custom
    ' 3ro. caracter
    '      1. Existe
    '      0. No Existe
       
     ' **************************************************************
     ' El Usuario no es mas usuario...
     ' **************************************************************
     If Mid$(NodoKey, 3, 1) = "0" Then
      ' Muestra: El usuario [ % ] ya no es mas usuario del Sistema...
      MostrarMSGBox MensajeRecurso(174) & Trim(Mid$(NodoKey, 4)) & MensajeRecurso(175), vbOKOnly, "vbInformation", Configuracion.TituloVentanas, True
      Exit Sub
     End If
     
     Select Case Mid$(NodoKey, 2, 1)
      Case 0:
       ' **************************************************************
       ' El Usuario no esta Conectado...
       ' **************************************************************
       ' Primero verifica si existe un Mensaje Offline Abierto...
       Respuesta = Varios.BuscarVentanaMensajeOffLine(Trim(Mid$(NodoKey, 4)))
       ' Si existe la muestra, Sino...
       If Respuesta <> 0 Then
        Forms(Respuesta).SetFocus
        Exit Sub
       End If
       ' Sino pregunta si quiere mandar un Menasaje Offline...
       Respuesta = MostrarMSGBox(MensajeRecurso(174) & Trim(Mid$(NodoKey, 4)) & MensajeRecurso(176) & MensajeRecurso(459), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
       If Respuesta = vbYes Then
        MenuClickAmigo_Click 8, "" ' Envia Mensaje OffLine - Metodo Estandar...
       End If
      Case 1:
       ' **************************************************************
       ' Abre la Ventana de Mensajes llamando al Metodo Estandar... (Como si
       ' clickeara el Menu emergente)
       ' **************************************************************
       MenuClickAmigo_Click 0, ""
      Case 2:
       ' **************************************************************
       ' El Usuario no esta Disponible...
       ' **************************************************************
       ' Primero verifica si existe un Mensaje Offline Abierto...
       Respuesta = Varios.BuscarVentanaMensajeOffLine(Trim(Mid$(NodoKey, 4)))
       ' Si existe la muestra, Sino...
       If Respuesta <> 0 Then
        Forms(Respuesta).SetFocus
        Exit Sub
       End If
       ' Sino pregunta si quiere mandar un Menasaje Offline...
       Respuesta = MostrarMSGBox(MensajeRecurso(174) & Trim(Mid$(NodoKey, 4)) & MensajeRecurso(177) & MensajeRecurso(459), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
       If Respuesta = vbYes Then
        MenuClickAmigo_Click 8, "" ' Envia Mensaje OffLine - Metodo Estandar...
       End If
      Case 3:
       ' **************************************************************
       ' Abre la Ventana de Mensajes llamando al Metodo Estandar... (Como si
       ' clickeara el Menu emergente)
       ' **************************************************************
       MenuClickAmigo_Click 0, ""
     End Select
  End If

SalirError:
 Exit Sub
 
End Sub
Private Sub MenuBotonConectar_Click()
 
 ' **************************************************************
 ' Ejecutar Sonido...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Menu Conectar...
 ' **************************************************************
 MenuBotonConectar.BackStyle = 1
 MenuBotonConectar.BackColor = Variables.FontFOndoMenuDescolgable
 MenuBotonConectar.ForeColor = Variables.FontMenuDescolgableAbierto
 MenuBotonConectar.BorderStyle = 1
 MenuToolConeccion.DesdeMenu = "Coneccion"
 MenuToolConeccion.ShowMenu EspacioIzquierda + Me.MenuBotonConectar.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonConectar.WidtH
  
End Sub
Private Sub MenuBotonConectar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Menu Conectar...
 ' **************************************************************
 If MenuToolConfiguracion.Visible Or _
    MenuToolAmigos.Visible Or _
    MenuToolAyuda.Visible Then
  Audio.EjecutarSonido "003"
  MenuBotonConectar.BackStyle = 1
  MenuBotonConectar.BackColor = Variables.FontFOndoMenuDescolgable
  MenuBotonConectar.ForeColor = Variables.FontMenuDescolgableAbierto
  MenuBotonConectar.BorderStyle = 1
  MenuToolConeccion.DesdeMenu = "Coneccion"
  MenuToolConeccion.ShowMenu EspacioIzquierda + Me.MenuBotonConectar.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonConectar.WidtH
 End If
 
End Sub
Private Sub MenuBotonConfiguracion_Click()
 
 ' **************************************************************
 ' Ejecutar Sonido...
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Menu Configuracion...
 ' **************************************************************
 MenuBotonConfiguracion.BackStyle = 1
 MenuBotonConfiguracion.BackColor = Variables.FontFOndoMenuDescolgable
 MenuBotonConfiguracion.ForeColor = Variables.FontMenuDescolgableAbierto
 MenuBotonConfiguracion.BorderStyle = 1
 MenuToolConfiguracion.DesdeMenu = "Configuracion"
 MenuToolConfiguracion.ShowMenu EspacioIzquierda + Me.MenuBotonConfiguracion.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonConfiguracion.WidtH
 
End Sub
Private Sub MenuBotonConfiguracion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Menu Configuracion...
 ' **************************************************************
 If MenuToolConeccion.Visible Or _
    MenuToolAmigos.Visible Or _
    MenuToolAyuda.Visible Then
   Audio.EjecutarSonido "003"
   MenuBotonConfiguracion.BackStyle = 1
   MenuBotonConfiguracion.BackColor = Variables.FontFOndoMenuDescolgable
   MenuBotonConfiguracion.ForeColor = Variables.FontMenuDescolgableAbierto
   'MenuBotonConfiguracion.BackColor = &HE0E0E0
   MenuBotonConfiguracion.BorderStyle = 1
   MenuToolConfiguracion.DesdeMenu = "Configuracion"
   MenuToolConfiguracion.ShowMenu EspacioIzquierda + Me.MenuBotonConfiguracion.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonConfiguracion.WidtH
 End If
 
End Sub
Private Sub MenuBotonAmigos_Click()

 ' **************************************************************
 ' Ejecutar Sonido...
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Menu Amigos...
 ' **************************************************************
 MenuBotonAmigos.BackStyle = 1
 MenuBotonAmigos.BackColor = Variables.FontFOndoMenuDescolgable
 MenuBotonAmigos.ForeColor = Variables.FontMenuDescolgableAbierto
 MenuBotonAmigos.BorderStyle = 1
 MenuToolAmigos.DesdeMenu = "Amigos"
 MenuToolAmigos.ShowMenu EspacioIzquierda + Me.MenuBotonAmigos.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonAmigos.WidtH

End Sub
Private Sub MenuBotonAmigos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Menu Amigos...
 ' **************************************************************
 If MenuToolConeccion.Visible Or _
    MenuToolConfiguracion.Visible Or _
    MenuToolAyuda.Visible Then
  Audio.EjecutarSonido "003"
  MenuBotonAmigos.BackStyle = 1
  MenuBotonAmigos.BackColor = Variables.FontFOndoMenuDescolgable
  MenuBotonAmigos.ForeColor = Variables.FontMenuDescolgableAbierto
  MenuBotonAmigos.BorderStyle = 1
  MenuToolAmigos.DesdeMenu = "Amigos"
  MenuToolAmigos.ShowMenu EspacioIzquierda + Me.MenuBotonAmigos.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonAmigos.WidtH
 End If
 
End Sub
Private Sub MenuBotonAyuda_Click()

 ' **************************************************************
 ' Ejecutar Sonido...
 ' **************************************************************
 Audio.EjecutarSonido "003"

 ' **************************************************************
 ' Menu Ayuda...
 ' **************************************************************
 MenuBotonAyuda.BackStyle = 1
 MenuBotonAyuda.BackColor = Variables.FontFOndoMenuDescolgable
 MenuBotonAyuda.ForeColor = Variables.FontMenuDescolgableAbierto
 MenuBotonAyuda.BorderStyle = 1
 MenuToolAyuda.DesdeMenu = "Ayuda"
 MenuToolAyuda.ShowMenu EspacioIzquierda + Me.MenuBotonAyuda.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonAyuda.WidtH
 
End Sub
Private Sub MenuBotonAyuda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 ' **************************************************************
 ' Menu Ayuda...
 ' **************************************************************
 If MenuToolConeccion.Visible Or _
    MenuToolConfiguracion.Visible Or _
    MenuToolAmigos.Visible Then
  Audio.EjecutarSonido "003"
  MenuBotonAyuda.BackStyle = 1
  MenuBotonAyuda.BackColor = Variables.FontFOndoMenuDescolgable
  MenuBotonAyuda.ForeColor = Variables.FontMenuDescolgableAbierto
  MenuBotonAyuda.BorderStyle = 1
  MenuToolAyuda.DesdeMenu = "Ayuda"
  MenuToolAyuda.ShowMenu EspacioIzquierda + Me.MenuBotonAyuda.Left + Cliente.Left, Cliente.Top + EspacioArriba, MenuBotonAyuda.WidtH
 End If
 
End Sub
Private Sub MenuClickUsuarioDesconectadoConectando_Click(ByVal Index As Long, Tag As String)
Dim Respuesta As Integer

 Select Case Index
  Case 0: ' Conectar
   ' **************************************************************
   ' Si el Usuario Esta Logueado le avisa que se desconectara su
   ' coneccion...
   ' **************************************************************
   If Configuracion.Logueado = 3 Then
    ' Muestra: Esta Acción lo Desconectara del Sistema...¿Desea Continuar?
    Respuesta = MostrarMSGBox(MensajeRecurso(178), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
    If Respuesta = vbNo Then
     Exit Sub
    End If
   End If
   ' **************************************************************
   ' Verifica si el Usuario esta logueado, si no lo esta, habre
   ' la ventana de Loguin...
   ' **************************************************************
   If Configuracion.Logueado <> 3 Then
    Load Loguin
    Variables.FormularioLoguin = True
    Variables.BringWindowToTop (Loguin.hwnd)
    Loguin.Show vbModal
    ' Loguin.SetFocus ' En Modal no es necesario
    Exit Sub
   End If
  Case 1: ' Preferencias
   Load Preferencias
   Preferencias.Show
  Case 2: ' Salir
   ' Muestra: ¿Está Seguro que Desea Salir del Sistema...?
   Respuesta = MostrarMSGBox(MensajeRecurso(179), vbYesNo, "vbQuestion", Configuracion.TituloVentanas)
   ' Cancela la salida del Sistema
   If Respuesta = vbNo Then
    Exit Sub
   End If
   ' Cierra Todo
   Cliente.TCPSocket.Close
   End ' TERMINA !!!!
 End Select

End Sub
Sub MenuMensajesPendientes_Click(ByVal Index As Long, Tag As String)
Dim Respuesta, Contador, ContadorTemp As Integer
Dim MensajeID As Integer
Dim Usuario, Mensaje As String
Dim HoraYFecha As String
Dim MensajesPendientesTemp() As Variables.MiMensajesPendientes

 ' **************************************************************
 ' No hay Mensajes...
 ' **************************************************************
 If Variables.CantidadDeMensajesPendientes = 0 Then Exit Sub

 ' **************************************************************
 ' Redefine el Repositorio de MensajesPendientes...
 ' **************************************************************
 ReDim MensajesPendientesTemp(Variables.CantidadDeMensajesPendientes)
 
 ' **************************************************************
 ' Suma uno ya que index empieza de 0, mientras que
 ' mensajes pendientes de 1
 ' **************************************************************
 MensajeID = Index + 1
 
 ' **************************************************************
 ' Busca la Ventana de Referencia
 ' **************************************************************
 Usuario = Trim(CStr(Variables.MensajesPendientes(MensajeID).MensajeDe))
 Respuesta = BuscarVentana(CStr(Usuario))
 If Respuesta = 0 Then
   Dim Handle As Long
   Handle = CrearVentanaMensaje(CStr(Usuario))
   Respuesta = BuscarVentanaHandle(Handle) ' BuscarVentana(CStr(Usuario))
 End If
 
 ' **************************************************************
 ' Pone el Aviso que es un mensaje Pendiente...
 ' **************************************************************
 Forms(Respuesta).AgregarLineaGrisConDatos 0, MensajeRecurso(463) & Usuario & "]..."
 
 ' **************************************************************
 ' Carga los Datos en el Formulario
 ' **************************************************************
 For Contador = 1 To Variables.CantidadDeMensajesPendientes
  If UCase(Trim(Variables.MensajesPendientes(Contador).MensajeDe)) = UCase(Usuario) Then
   Mensaje = Trim(CStr(Variables.MensajesPendientes(Contador).Mensaje))
   HoraYFecha = Trim(CStr(Variables.MensajesPendientes(Contador).HoraYFecha))
   Forms(Respuesta).AgregarMensaje CStr(Usuario), CStr(Mensaje), CStr(HoraYFecha)
   Variables.MensajesPendientes(Contador).Mensaje = ""
  End If
 Next
 
 ' **************************************************************
 ' Borra el Mensajes Pendiente
 ' **************************************************************
 ' Carga los Mensajes que quedaron...
 ContadorTemp = 0
 For Contador = 1 To Variables.CantidadDeMensajesPendientes
  If Trim(Variables.MensajesPendientes(Contador).Mensaje) <> "" Then
   ContadorTemp = ContadorTemp + 1
   MensajesPendientesTemp(ContadorTemp) = Variables.MensajesPendientes(Contador)
  End If
 Next
 ' Pasa y define los Mensajes que quedan...
 For Contador = 1 To ContadorTemp
  Variables.MensajesPendientes(Contador) = MensajesPendientesTemp(Contador)
 Next
 ' Termina de Definir los Mensaje
 Varios.CambiarMensajesPendientes (ContadorTemp)
 
 ' **************************************************************
 ' Graba los Mensajes Pendientes
 ' **************************************************************
 GrabarMensajesPendientes
 
End Sub
Private Sub MenuPopUpTray_Click(ByVal Index As Long, Tag As String)
 
 Select Case Index
  Case 0: ' Conectar
   MenuToolConeccion_Click 0, ""
  Case 1: ' Desconectar
   MenuToolConeccion_Click 1, ""
  Case 2: ' Cambiar Usuario...
   MenuToolConeccion_Click 2, ""
  Case 3: ' Preferencias
   MenuToolConfiguracion_Click 5, ""
  Case 5: ' Mostrar EIM
   MenuPopUpTray.Hide
   Me.WindowState = 0
   ' Muestra la Ventana arriba de las Demas
   If Variables.FormularioLoguin = False Then
     Varios.DefinirPosicionDeCliente True
     Varios.DefinirPosicionDeCliente
     Variables.BringWindowToTop (Me.hwnd)
     ' Solo hace un SetFocus si esta visible (Evita Errores)
     If Me.Visible Then
      Me.SetFocus
     End If
    Else
     Variables.BringWindowToTop (Loguin.hwnd)
     ' Solo hace un SetFocus si esta visible (Evita Errores)
     If Loguin.Visible Then
      Loguin.SetFocus
     End If
   End If
  Case 6: ' Sobre  EIM
   MenuToolAyuda_Click 2, ""
  Case 8: ' Mostrar Estados...
   If Tag = "PerdioFoco" Then
    MenuCambioDeEstado.Hide
    MenuPopUpTray.SeLlamoAlDesplegable = False
    Exit Sub
   End If
   ' Solo carga si no esta visible...
   If MenuCambioDeEstado.Visible = True Then Exit Sub
   ' Carga el Menu de Grupos
   Set MenuCambioDeEstado = New IcoMenu
   With MenuCambioDeEstado
    ' Disponible (Normal)...
    .SetItem 0, MensajeRecurso(180), Imagenes.ListImages("EstadoVisible").Picture
    ' No Disponible...
    .SetItem 1, MensajeRecurso(181), Imagenes.ListImages("EstadoNoDisponible").Picture
    ' Enseguida Vuelvo...
    .SetItem 2, MensajeRecurso(182), Imagenes.ListImages("EstadoCustom").Picture
    ' No Molestar...
    .SetItem 3, MensajeRecurso(183), Imagenes.ListImages("EstadoCustom").Picture
    ' Definir Estado... (Custom)
    .SetItem 4, MensajeRecurso(184), Imagenes.ListImages("EstadoCustom").Picture
   End With
    MenuCambioDeEstado.ShowMenu MenuPopUpTray.Left - MenuCambioDeEstado.WidtH + 900, MenuPopUpTray.Top + MenuPopUpTray.LBLmenu(Index).Top + MenuPopUpTray.LBLmenu(Index).HeighT - MenuCambioDeEstado.HeighT + 2890, , MenuPopUpTray.hwnd
  Case 10: ' Salir
   MenuToolConeccion_Click 4, ""
 End Select
 
End Sub
Private Sub MenuCambioDeEstado_Click(ByVal Index As Long, Tag As String)

  ' **************************************************************
  ' Llama al Menu para el Cambio de Estado...
  ' **************************************************************
  Me.MenuDeEstadosDeUsuario_Click Index, Tag
  
End Sub
Private Sub MenuClickUsuario_Click(ByVal Index As Long, Tag As String)
Dim NombreGrupo As String
 
 Select Case Index
  Case 0: ' Cambiar Mis Datos
   Varios.VerDatosUsuario Trim(Configuracion.IDAliasUsuario), True
  Case 1: ' Ver Mis Datos
   Varios.VerDatosUsuario Trim(Configuracion.IDAliasUsuario), False
  Case 2: ' Agregar Amigos
   Varios.RecargarListadoDeAmigos
  Case 4: ' Crear Grupo
   Varios.CrearGrupo
  Case 5: ' Eliminar Grupo
   Load EliminarGrupoOAmigo
   EliminarGrupoOAmigo.MostrarFormulario ("Grupo")
   ' Verifica que Exista algo Que Borrar
   If EliminarGrupoOAmigo.CantidadActual = 0 Then
    ' No Existen Grupos que Pueda Borrar...
    MostrarMSGBox MensajeRecurso(185), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    Unload EliminarGrupoOAmigo ' Descarga el Form...
    Exit Sub ' Sale sin Hacer Nada...
   End If
   EliminarGrupoOAmigo.Show ' vbModal
  Case 7: ' Agregar Amigo
   Varios.AgregarAmigo ("")
  Case 8: ' Eliminar Amigo
   Load EliminarGrupoOAmigo
   EliminarGrupoOAmigo.MostrarFormulario ("Amigo")
   ' Verifica que Exista algo Que Borrar
   If EliminarGrupoOAmigo.CantidadActual = 0 Then
    ' No Existen Amigos que Pueda Borrar...
    MostrarMSGBox MensajeRecurso(186), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    'Unload EliminarGrupoOAmigo
    'Exit Sub
   End If
   EliminarGrupoOAmigo.Show ' vbModal
  Case 9: ' Buscar Amigo
   Varios.AgregarAmigo ("")
  Case 10: ' Bloqueo de Amigos
   Load UsuariosBloqueados
   UsuariosBloqueados.Show
  Case 12: ' Preferencias
   Load Preferencias
   Preferencias.Show
 End Select
 
End Sub
Private Sub MenuClickGrupo_Click(ByVal Index As Long, Tag As String)
Dim NombreGrupo As String
 
 Select Case Index
  'Case 0: ' Enviar Mensaje Grupal
   ' **************************************************************
   ' No implementado...
   ' **************************************************************
   ' Opción No Implementada...
   'MostrarMSGBox MensajeRecurso(187), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Case 0: ' Eliminar Grupo
   Varios.EliminarGrupo Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 2)
  Case 1: ' Renombrar Grupo
   Varios.RenombrarGrupo Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 2)
  Case 3: ' Agregar Amigo
   ' **************************************************************
   ' Definir en que grupo se debe crear el Usuario
   ' **************************************************************
   NombreGrupo = Trim(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 2))
   ' **************************************************************
   ' Crea el Amigo
   ' **************************************************************
   Varios.AgregarAmigo (NombreGrupo)
 End Select
 
End Sub
Private Sub MenuClickAmigo_Click(ByVal Index As Long, Tag As String)
Dim Amigo As String
Dim Mensajeria As New Mensajes
Dim Respuesta As Integer

 Select Case Index
  Case 0: ' Enviar Mensaje
   Respuesta = BuscarVentana(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4))
   ' Verifica si existe la Ventana si Existe, se
   ' posiciona en ella, sino abre una nueva...
   If Respuesta = 0 Then
     CrearVentanaMensaje Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4)
    Else
     Forms(Respuesta).Show
     If Forms(Respuesta).Visible Then
      Forms(Respuesta).SetFocus
     End If
   End If
  'Case 1: ' Enviar Archivo
   ' **************************************************************
   ' No implementado...
   ' **************************************************************
   ' Opción No Implementada...
  ' MostrarMSGBox MensajeRecurso(187), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Case 2: ' Ver Datos del Amigo
   Amigo = Trim(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4))
   Varios.VerDatosUsuario Amigo, False
  Case 4: ' Elimnar Amigo
    Varios.EliminarAmigo Trim(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4))
  Case 6: ' Mover a Grupo
   ' Si perdio el Foco lo Descarga...
   If Tag = "PerdioFoco" Then
    MenuDeCambioDeGrupo.Hide
    MenuClickAmigo.SeLlamoAlDesplegable = False
    Exit Sub
   End If
   ' Solo carga si no esta visible...
   If MenuDeCambioDeGrupo.Visible = True Then Exit Sub
   ' Carga el Menu de Grupos
   Set MenuDeCambioDeGrupo = New IcoMenu
   With MenuDeCambioDeGrupo
    Dim GrupoDelAmigo As String
    GrupoDelAmigo = Trim(Varios.BuscarElGrupoDelAmigo(Trim(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4))))
    ' Fuera
    If GrupoDelAmigo = "" Then GrupoDelAmigo = MensajeRecurso(148)
    ' Ver Más...
    .SetItem 0, MensajeRecurso(188), ImagenesMenus.ListImages("MoverAGrupo").Picture, CompletarCadena(GrupoDelAmigo, 20, "I", " ") & CompletarCadena(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4), 16, "I", " ")
    .SetItem 1, ""
    ' --> Ojo que al cambiar aca hay que cambiar en Cargar Menus... ' FUERA
    ' Fuera de Grupo...
    .SetItem 2, MensajeRecurso(132), ImagenesMenus.ListImages("MoverAGrupo").Picture, CompletarCadena(MensajeRecurso(148), 20, "I", " ") & CompletarCadena(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4), 16, "I", " ")
    ' **************************************************************
    ' Carga los Grupos
    ' **************************************************************
    Dim Contador, Cantidad As Integer
    Cantidad = 0
    For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
     If Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1) = "G" Then
      Cantidad = Cantidad + 1
      ' --> Ojo que al cambiar aca hay que cambiar en Cargar Menus...
      .SetItem 2 + Cantidad, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2), ImagenesMenus.ListImages("MoverAGrupo").Picture, CompletarCadena(Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2), 20, "I", " ") & CompletarCadena(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4), 16, "I", " ")
     End If
    Next
   End With
   MenuDeCambioDeGrupo.ShowMenu MenuClickAmigo.Left + MenuClickAmigo.WidtH, MenuClickAmigo.LBLmenu(Index).Top + MenuClickAmigo.Top - 50, , MenuClickAmigo.hwnd
  Case 8: ' Enviar Mensaje Off-Line...
   Respuesta = Varios.BuscarVentanaMensajeOffLine(Trim(Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4)))
   If Respuesta = 0 Then
     CrearVentanaMensajeOffLine Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4), MensajeRecurso(454) & Mid$(Cliente.ListadoDeAmigos.SelectedItem.key, 4) & "]..."
    Else
     If Forms(Respuesta).Visible Then
       Forms(Respuesta).SetFocus
      Else
       Forms(Respuesta).Show
     End If
   End If
 End Select
  
End Sub
Private Sub MenuDeCambioDeGrupo_Click(ByVal Index As Long, Tag As String)
Dim UsuarioID As String
Dim Grupo As String

 Grupo = Trim(Mid$(Tag, 1, 20))
 UsuarioID = Trim(Mid$(Tag, 21))

 Select Case Index
  Case 0: ' Muestra la Ventana de Cambio de grupo
   Load CambiarAGrupo
   CambiarAGrupo.MostrarFormulario Grupo, UsuarioID
  Case Else
   Varios.CambiarUsuarioDeGrupo UsuarioID, Grupo, False
 End Select
 
End Sub
Private Sub MenuToolAyuda_Click(ByVal Index As Long, Tag As String)

 Select Case Index
  Case 0:
   ' **************************************************************
   ' No implementado...
   ' **************************************************************
   ' Muestra: Opción No Implementada...
   MostrarMSGBox MensajeRecurso(187), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Case 2:
   ' **************************************************************
   ' About EIM
   ' **************************************************************
   Load Presentacion
   Presentacion.Touch.Visible = True
   Presentacion.TimeOut.Enabled = False
   Presentacion.Show
 End Select
  
End Sub
Private Sub MenuToolAmigos_Click(ByVal Index As Long, Tag As String)

 Select Case Index
  Case 0:
    Varios.RecargarListadoDeAmigos
  Case 2:
    Varios.CrearGrupo
  Case 3: ' Eliminar Grupo
   Load EliminarGrupoOAmigo
   EliminarGrupoOAmigo.MostrarFormulario ("Grupo")
   ' Verifica que Exista algo Que Borrar
   If EliminarGrupoOAmigo.CantidadActual = 0 Then
    ' No Existen Grupos que Pueda Borrar...
    MostrarMSGBox MensajeRecurso(185), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
    Unload EliminarGrupoOAmigo
    Exit Sub
   End If
   EliminarGrupoOAmigo.Show 'vbModal
  Case 5:
   Varios.AgregarAmigo ("")
  Case 6: ' Eliminar Amigos
   Load EliminarGrupoOAmigo
   EliminarGrupoOAmigo.MostrarFormulario ("Amigo")
   ' Verifica que Exista algo Que Borrar
   If EliminarGrupoOAmigo.CantidadActual = 0 Then
    ' No Existen Amigos que Pueda Borrar...
    MostrarMSGBox MensajeRecurso(186), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   End If
   EliminarGrupoOAmigo.Show 'vbModal
  Case 7:
   Varios.AgregarAmigo ("")
  Case 9:
   Load UsuariosBloqueados
   UsuariosBloqueados.Show
 End Select
  
End Sub
Private Sub MenuToolConfiguracion_Click(ByVal Index As Long, Tag As String)

 Select Case Index
  Case 0:
    VerDatosUsuario Trim(Configuracion.IDAliasUsuario), False
  Case 1:
    Varios.VerDatosUsuario Trim(Configuracion.IDAliasUsuario), True
  Case 3:
   ' **************************************************************
   ' Cambio de Password
   ' **************************************************************
   Load CambioDePassword
   CambioDePassword.Show 'vbModal
  Case 5:
   ' **************************************************************
   ' Preferencias
   ' **************************************************************
   Load Preferencias
   Preferencias.Show
 End Select
 
End Sub

Private Sub MenuToolConeccion_Click(ByVal Index As Long, Tag As String)
Dim Respuesta As Integer

 Select Case Index
  Case 0: ' Conectar
   ' **************************************************************
   ' Si el Usuario Esta Logueado le avisa que se desconectara su
   ' coneccion...
   ' **************************************************************
   If Configuracion.Logueado = 3 Then
    ' Esta Acción lo Desconectara del Sistema...¿Desea Continuar?
    Respuesta = MostrarMSGBox(MensajeRecurso(178), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
    If Respuesta = vbNo Then
     Exit Sub
    End If
   End If
   ' **************************************************************
   ' Verifica si el Usuario esta logueado, si no lo esta, habre
   ' la ventana de Loguin...
   ' **************************************************************
   If Configuracion.Logueado <> 3 Then
    Load Loguin
    Variables.FormularioLoguin = True
    Variables.BringWindowToTop (Loguin.hwnd)
    Loguin.Show vbModal
    Exit Sub
   End If
  Case 1: ' Desconectar
   ' ¿Está Seguro que Desea Desconectarse?...
   Respuesta = MostrarMSGBox(MensajeRecurso(192) & Chr$(13) & MensajeRecurso(469), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
   If Respuesta = vbNo Then
    Exit Sub
   End If
   ' **************************************************************
   ' Desconectarse del Servidor
   ' **************************************************************
   ' Desconecta al Cliente
   TCPSocket.Close
   ' Pasa a Estado Desconectado
   SocketTCP.CambiarEstadoDelCliente (0)
  Case 2: ' Cambiar Usuario
   ' **************************************************************
   ' Si el Usuario Esta Logueado le avisa que se desconectara su
   ' coneccion...
   ' **************************************************************
   If Configuracion.Logueado = 3 Then
    ' Esta Acción lo Desconectara del Sistema...¿Desea Continuar?
    Respuesta = MostrarMSGBox(MensajeRecurso(178), vbYesNo, "vbInformation", Configuracion.TituloVentanas)
    If Respuesta = vbNo Then
     Exit Sub
    End If
   End If
   ' **************************************************************
   ' Desconectarse del Servidor
   ' **************************************************************
   ' Desconecta al Cliente
   TCPSocket.Close
   ' Pasa a Estado Desconectado
   SocketTCP.CambiarEstadoDelCliente (0)
   ' **************************************************************
   ' Verifica si el Usuario esta logueado, si no lo esta, habre
   ' la ventana de Loguin...
   ' **************************************************************
   If Configuracion.Logueado <> 3 Then
    Load Loguin
    Variables.FormularioLoguin = True
    Variables.LoguinAutomatico = False
    Variables.BringWindowToTop (Loguin.hwnd)
    Loguin.Show vbModal
    Exit Sub
   End If
  Case 4: ' Salir
   Varios.SalirDelSistema
 End Select
 
End Sub
Private Sub ContextualAgregarAmigo_Click()

 ' **************************************************************
 ' Menu Agregar Amigos
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Varios.AgregarAmigo ("")
 
End Sub
Private Sub ContextualAgregarAyuda_Click()

 ' **************************************************************
 ' Menu Agregar Amigos
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Varios.AgregarAmigo ("")
 
End Sub
Private Sub ContextualAgregarGrupo_Click()

 ' **************************************************************
 ' Menu Agregar Grupo
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Varios.CrearGrupo

End Sub
Private Sub ContextualRefrescarAmigos_Click()

 ' **************************************************************
 ' Menu Refrescar Amgigos
 ' **************************************************************
 Audio.EjecutarSonido "003"
 Varios.RecargarListadoDeAmigos
 
End Sub

Private Sub ImageMenuDeEstado_Click()

 ' **************************************************************
 ' Menu Menu de Estado
 ' **************************************************************
 Audio.EjecutarSonido "003"
 MenuDeEstadosDeUsuario.ShowMenu Me.ImageMenuDeEstado.Left + Cliente.Left - 20, Cliente.Top + 1565, 690

End Sub
Public Sub MenuDeEstadosDeUsuario_Click(ByVal Index As Long, Tag As String)
Dim Numero, Texto As String

 ' **************************************************************
 ' Ojo que al tocar este modulo afecta directamente a
 ' MenuCambioDeEstado_Click, ya que el modulo antes mencionado
 ' cambia el estado llamado a Menudeestadosusuario_click
 ' **************************************************************

 ' **************************************************************
 ' Verifica Que Numero se Clickeo
 ' **************************************************************
 Select Case Index
  Case 0
   Numero = "1"
   Texto = ""
  Case 1
   Numero = "2"
   Texto = ""
  Case 4
   Numero = "3"
   ' Por Favor Ingrese el Texto de su Estado:
   Texto = Trim(MostrarInputBox(MensajeRecurso(193), 20, Configuracion.TituloVentanas))
   ' Verifica que el Estado Ingresado sea Valido...
   ' Sino Sale...
   If Texto = "" Then
    Exit Sub
   End If
  Case Else
   Numero = "3"
   Texto = Trim(MenuDeEstadosDeUsuario.LBLmenu(Index))
 End Select
 
 ' **************************************************************
 ' Pone Localmente el Estado (En una variable... Lo pone
 ' realmente cuando el servidor confirma el Cambio de estado...)
 ' **************************************************************
 NuevoEstadoUsuario.Numero = CInt(Numero)
 NuevoEstadoUsuario.Texto = Texto ' Lo trimea(TRIM) mas arriba...
 
 ' **************************************************************
 ' Envia el Nuevo Estado
 ' **************************************************************
 SocketTCP.EnviarCambioDeEstado NuevoEstadoUsuario.Numero, NuevoEstadoUsuario.Texto

End Sub
Private Sub Form_Load()
Dim Posicion As POINTAPI

 ' **************************************************************
 ' Define la Posicion del Form
 ' **************************************************************
 Posicion = Inicializar.PosicionCliente("Leer", Me.Left, Me.Top)
 If Posicion.X <> -10000 Then
  Me.Left = Posicion.X
 End If
 If Posicion.Y <> -10000 Then
  Me.Top = Posicion.Y
 End If
 
 ' **************************************************************
 ' Carga el Icono de Aplicacion
 ' **************************************************************
 Me.CargarTextos
 Me.Icon = Cliente.Icon
 Me.FormularioNombre = "Cliente"
 
 ' **************************************************************
 ' Setea las Ventanas
 ' **************************************************************
 InicializarSistema
 Me.Caption = Trim(Configuracion.TituloVentanas)
   
 ' **************************************************************
 ' Carga el Estado del Icono de Sonido
 ' **************************************************************
 DefinirEstadoSonido
  
 ' **************************************************************
 ' Cambia el Color del Listado de Amigos
 ' **************************************************************
 Call SendMessage(Cliente.hwnd, 4381&, 0, vbWhite)
  
 ' **************************************************************
 ' Cargar los Menus
 ' **************************************************************
 CargarLosMenus
 
 ' **************************************************************
 ' Verifica si el Usuario esta logueado, si no lo esta, habre
 ' la ventana de Loguin... (En modo VBModal)
 ' **************************************************************
 If Configuracion.Logueado <> 3 Then
  Load Loguin
  Variables.FormularioLoguin = True
  Loguin.Show vbModal
  Variables.BringWindowToTop (Loguin.hwnd)
 End If
 
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msgCallBackMessage, Resultado As Long
Dim PosicionY, PosicionX, TamanioX, TamanioY  As Single
 
 ' **************************************************************
 ' Tomar la posicion del Area de Trabajo
 ' **************************************************************
 Resultado = SystemParametersInfo(SPI_GETWORKAREA, 0&, Variables.AreaDeTrabajo, 0&)
 PosicionY = AreaDeTrabajo.Bottom
 PosicionY = PosicionY * Screen.TwipsPerPixelY
 PosicionX = AreaDeTrabajo.Right
 PosicionX = PosicionX * Screen.TwipsPerPixelX
 
 On Error GoTo ErrorOnFocus
  
  ' **************************************************************
  ' Aca se Maneja las Opciones cuando con el Mouse
  ' Se hace click sobre el Tray Icon
  ' **************************************************************
  msgCallBackMessage = X / Screen.TwipsPerPixelX
   
  ' **************************************************************
  ' Actua segun la Accion del Muose
  ' **************************************************************
  Select Case msgCallBackMessage
    ' Doble Click
    Case WM_LBUTTONDBLCLK
     ' Muestra el Cliente
     MenuPopUpTray.Hide
     Me.WindowState = 0
     ' Muestra la Ventana arriba de las Demas
     If Variables.FormularioLoguin = False Then
       Varios.DefinirPosicionDeCliente True
       Me.SetFocus
       Variables.BringWindowToTop (Me.hwnd) ' On-Top
      Else
       Varios.CerrarVentanasDeMenus
       Variables.BringWindowToTop (Loguin.hwnd) ' On-Top
       Loguin.Show vbModal
     End If
     
    ' Boton Izquierdo Sobre el Icon TRAY
    Case WM_LBUTTONDOWN
     MenuPopUpTray.Hide
     Me.WindowState = 0
     ' Muestra la Ventana arriba de las Demas
     If Variables.FormularioLoguin = False Then
       Varios.DefinirPosicionDeCliente True
       Variables.BringWindowToTop (Me.hwnd) ' On-Top
       If Configuracion.Logueado <> 3 Then
        Varios.CerrarVentanasDeMenus
        Load Loguin
        Variables.FormularioLoguin = True
        Variables.BringWindowToTop (Loguin.hwnd) ' On-Top
        Loguin.Show vbModal
       End If
      Else
       Varios.CerrarVentanasDeMenus
       Variables.BringWindowToTop (Loguin.hwnd) ' On-Top
       Loguin.Show vbModal
     End If
    
    ' Boton Derecho sobre el Icon TRAY
    Case WM_RBUTTONDOWN
     MenuPopUpTray.Hide
     If Variables.FormularioLoguin = True Then
       Varios.CerrarVentanasDeMenus
       Variables.BringWindowToTop (Loguin.hwnd)  ' On-Top
       Loguin.Show vbModal
       Exit Sub
     End If
     TamanioY = 270 * MenuPopUpTray.LBLmenu.Count + 90
     TamanioX = MenuPopUpTray.MaxWidth + 90
    ' Define que Muestra cuando esta o No Conectado en el PopUpTray
     Select Case Configuracion.Logueado
      Case 0 ' Desconectado
       Cliente.MenuPopUpTray.HabilitarItem 0, True
       Cliente.MenuPopUpTray.HabilitarItem 1, False
       Cliente.MenuPopUpTray.HabilitarItem 8, False
      Case 1 ' Conectando
       Cliente.MenuPopUpTray.HabilitarItem 0, False
       Cliente.MenuPopUpTray.HabilitarItem 1, False
       Cliente.MenuPopUpTray.HabilitarItem 8, False
      Case 3 ' Conectado
       Cliente.MenuPopUpTray.HabilitarItem 0, False
       Cliente.MenuPopUpTray.HabilitarItem 1, True
       Cliente.MenuPopUpTray.HabilitarItem 8, True
     End Select
     MenuPopUpTray.ShowMenu PosicionX - TamanioX - 10, PosicionY - TamanioY - 10
     
  End Select

Salir:
 Exit Sub
 
ErrorOnFocus:
 Resume Salir
 
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 ' **************************************************************
 ' Descarga el Icono del Sistem Tray
 ' **************************************************************
 SysIcon.HideIcon

End Sub
Private Sub Form_Unload(Cancel As Integer)

 ' **************************************************************
 ' Verifica el Paquete
 ' **************************************************************
 SysIcon.HideIcon
 
End Sub
Private Sub ListadoDeAmigos_NodeClick(ByVal Node As MSComctlLib.Node)
Dim Posicion As POINTAPI
Dim Contador, Cantidad, PosicionCaracter As Integer
Dim ItemMenu, HoraYFecha As String
 
 ' **************************************************************
 ' Ejecuta el Sonido
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Toma la Posicion del Mouse
 ' **************************************************************
 GetCursorPos Posicion
 
 ' **************************************************************
 ' Dispara los Mensajes Pendientes
 ' **************************************************************
 If Mid$(Node.Text, 1, 1) = "*" Then
  Select Case Variables.CantidadDeMensajesPendientes
   Case 0
    ' **************************************************************
    ' No posee mensaje pendientes
    ' **************************************************************
    ' Usted no Posee Mensajes Pendientes...
    MostrarMSGBox MensajeRecurso(194), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
   Case Else
    ' Carga el Menu de Pendientes
    CargarMenuMensajesPendientes
    ' Muestra el Menu
    MenuMensajesPendientes.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
   End Select
 End If
 
 ' **************************************************************
 ' Abre el PopUp del Usuario
 ' **************************************************************
 If UCase(Mid$(Node.key, 1, 7)) = "USUARIO" Then
  Select Case Configuracion.Logueado
   Case 0 ' Desconectado
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 0, True
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 1, True
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 2, True
    Cliente.MenuClickUsuarioDesconectadoConectando.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
   Case 1 ' Conectando
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 0, False
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 1, True
    Cliente.MenuClickUsuarioDesconectadoConectando.HabilitarItem 2, True
    Cliente.MenuClickUsuarioDesconectadoConectando.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
   Case 3 ' Conectando
    MenuClickUsuario.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
  End Select
  Exit Sub
 End If
 
 ' **************************************************************
 ' Abre el PopUp de Amigo
 ' **************************************************************
 If Mid$(Node.key, 1, 1) = "U" And UCase(Mid$(Node.key, 1, 7)) <> "USUARIO" Then
  ' **************************************************************
  ' Define que muestra y que no... (Existe o No)
  ' **************************************************************
  ' Primero Habilita Todo
  MenuClickAmigo.HabilitarItem 0, True
  MenuClickAmigo.HabilitarItem 1, True
  MenuClickAmigo.HabilitarItem 3, True
  MenuClickAmigo.HabilitarItem 2, True
  MenuClickAmigo.HabilitarItem 8, True
  
  If Mid$(Node.key, 3, 1) = "1" Then
    ' **************************************************************
    ' Segun el Estado Continua Viendo que Muestra y Que No
    ' **************************************************************
    Select Case Mid$(Node.key, 2, 1)
     Case "0"
      MenuClickAmigo.HabilitarItem 0, False
      MenuClickAmigo.HabilitarItem 1, False
      MenuClickAmigo.HabilitarItem 8, True
     Case "1"
      
     Case "2"
      MenuClickAmigo.HabilitarItem 0, False
      MenuClickAmigo.HabilitarItem 1, False
      MenuClickAmigo.HabilitarItem 8, True
     Case "3"
      MenuClickAmigo.HabilitarItem 0, True
      MenuClickAmigo.HabilitarItem 1, True
      MenuClickAmigo.HabilitarItem 8, False
    End Select
   Else
    MenuClickAmigo.HabilitarItem 0, False
    MenuClickAmigo.HabilitarItem 1, False
    MenuClickAmigo.HabilitarItem 2, False
    MenuClickAmigo.HabilitarItem 3, False
    MenuClickAmigo.HabilitarItem 8, False
    
  End If
  ' **************************************************************
  ' Muestra el Menu
  Me.UltimoClickDeAmigo = Node.key
  MenuClickAmigo.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
  ' PopupMenu ClickAmigo
  Exit Sub
 End If
 
 ' **************************************************************
 ' Abre el PopUp de Grupo
 ' **************************************************************
 If Mid$(Node.key, 1, 1) = "G" Then
  MenuClickGrupo.ShowMenu Posicion.X * Screen.TwipsPerPixelX, Posicion.Y * Screen.TwipsPerPixelY
  Exit Sub
 End If
 
End Sub
Private Function BuscarFechaMensajesPendientes(UserID As String) As String
Dim Contador, Cantidad As Integer
Dim Fecha As String

 ' **************************************************************
 ' Busca la cantidad de Mensajes para el UserID
 ' **************************************************************
 Cantidad = 0
 For Contador = 1 To CantidadDeMensajesPendientes
  If Trim(UCase(MensajesPendientes(Contador).MensajeDe)) = Trim(UCase(UserID)) Then
   Fecha = MensajesPendientes(Contador).HoraYFecha
  End If
 Next
 
 ' **************************************************************
 ' Devuelve el Valor
 ' **************************************************************
 BuscarFechaMensajesPendientes = Fecha
 
End Function

Private Sub CargarMenuMensajesPendientes()
Dim Cantidad, CantidadMensajes, Contador, PosicionCaracter As Integer
Dim ItemMenu, HoraYFecha, Fecha As String

    ' **************************************************************
    ' Agrupa los Mensajes Pendientes
    ' **************************************************************
    Varios.AgruparMensajesPendientes
    
    ' **************************************************************
    ' Trabaja con Mensaje Pendientes
    ' **************************************************************
    If Variables.CantidadDeMensajesPendientesAgrupados > 8 Then
      Cantidad = 8
     Else
      Cantidad = Variables.CantidadDeMensajesPendientesAgrupados
    End If
    
    ' **************************************************************
    ' Carga el Menu con los Mensajes Pendientes
    ' **************************************************************
    Set MenuMensajesPendientes = New IcoMenu
    For Contador = 1 To CantidadDeMensajesPendientesAgrupados
     With MenuMensajesPendientes
      ' Segun la Cantidad de Mensajes define el texto...
      If Variables.MensajesPendientesAgrupados(Contador).Cantidad > 1 Then
        ' Varios Mensajes
        '  Mensajes)
        ItemMenu = Variables.MensajesPendientesAgrupados(Contador).MensajeDe & " (" & Variables.MensajesPendientesAgrupados(Contador).Cantidad & MensajeRecurso(195)
       Else
        Fecha = BuscarFechaMensajesPendientes(Variables.MensajesPendientes(Contador).MensajeDe)
        PosicionCaracter = InStr(Fecha, "_")
        HoraYFecha = Left(Fecha, PosicionCaracter - 1) & " " & Mid$(Fecha, PosicionCaracter + 1)
        ItemMenu = Variables.MensajesPendientesAgrupados(Contador).MensajeDe & " [" & HoraYFecha & "]"
      End If
      .SetItem Contador - 1, ItemMenu, ImagenesMenus.ListImages("Mensaje").Picture
     End With
    Next
  
End Sub
Private Sub OnLineTime_Timer()
Dim TiempoActual, Diferencia As String
Dim DiferenciaHH, DiferenciaMM, DiferenciaSS As Double

 If OnLineTime Then
  
  ' Calcula la Diferencia de Tiempo
  TiempoActual = Date & " " & Time
  DiferenciaSS = DateDiff("s", Variables.TiempoEnLineaContanteDesde, TiempoActual)
  DiferenciaMM = Int(DiferenciaSS / 60)
  DiferenciaHH = Int(DiferenciaSS / 3600)
  
  ' Resta los tiempos a las diferencia
  DiferenciaMM = DiferenciaMM - (DiferenciaHH * 60)
  DiferenciaSS = DiferenciaSS - (DiferenciaMM * 60) - (DiferenciaHH * 3600)
    
  If DiferenciaHH < 0 Then DiferenciaHH = DiferenciaHH * -1
  
  ' Armar la cadena de Diferencia
  Diferencia = Varios.CompletarCadena(CStr(DiferenciaHH), 2, "I", "0") & ":" & _
               Varios.CompletarCadena(CStr(DiferenciaMM), 2, "I", "0") & ":" & _
               Varios.CompletarCadena(CStr(DiferenciaSS), 2, "I", "0")
  
  If DiferenciaHH > 99 Then
   ' Demasiado
   Diferencia = MensajeRecurso(196)
  End If
  
  ' Poner el Tiempo De Diferencia
  Cliente.TiempoEnLinea = Diferencia
  
  ' **************************************************************
  ' Define si por la Inactividad lo debe pasa a 'Enseguida Vuelvo'
  ' **************************************************************
  Diferencia = DateDiff("s", Variables.UltimoMensajeEnviado, Time)
  If CLng(Diferencia) >= CLng((Configuracion.TiempoParaPasaraInactivo) * 60) Then
   ' Verifica que el estado actual sea conactado Visible
   If Configuracion.EstadoDelUsuario = 1 Then
    ' Lo Pasa a Enseguida vuelvo
    MenuDeEstadosDeUsuario_Click 2, ""
    ' Deinfe que paso a Enseguida Vuelvo
    Variables.PasoAEnseguidaVuelvo = True
   End If
  End If
  
 End If
  
End Sub

Private Sub RefrescoAmigos_Timer()

 ' Envia el Paquete de Refresco de Amigos
 If Configuracion.Logueado = 3 Then
  EnviarPaqueteTCP ("22")
 End If
 
End Sub
Private Sub ScrollAbajo_Click()
 
 ' Se Mueve de a dos...
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Abajo"
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Abajo"

End Sub

Private Sub ScrollArriba_Click()

 ' Se Mueve de a dos...
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Arriba"
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Arriba"
 
End Sub
Private Sub ScrollDerecha_Click()

  ' Se Mueve de a dos...
  Scroller.ScrollTreeView Me.ListadoDeAmigos, "Derecha"
  Scroller.ScrollTreeView Me.ListadoDeAmigos, "Derecha"
  
End Sub
Private Sub ScrollIzquierda_Click()

 ' Se Mueve de a dos...
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Izquierda"
 Scroller.ScrollTreeView Me.ListadoDeAmigos, "Izquierda"
 
End Sub
Private Sub SonidoSeteo_Click()

 ' **************************************************************
 ' Solo deja cambiarlo si esta logueado, ya que esta config se
 ' aplica al usuario...
 ' **************************************************************
 'If Configuracion.Logueado <> 3 Then Exit Sub
 
 If Configuracion.SonidoActivado Then
   Configuracion.SonidoActivado = False
   'SonidoSeteo.Picture = Me.ImagenesSonido.ListImages("SinSonido").Picture
   'Cliente.SonidoSeteo.ToolTipText = MensajeRecurso(315)
  Else
   Configuracion.SonidoActivado = True
   'SonidoSeteo.Picture = Me.ImagenesSonido.ListImages("ConSonido").Picture
   'Cliente.SonidoSeteo.ToolTipText = MensajeRecurso(314)
   ' **************************************************************
   ' Ejecuta el Sonido de Click
   ' **************************************************************
   Audio.EjecutarSonido "003"
 End If
 
 DefinirEstadoSonido
 
 Inicializar.GrabarConfiguracion
 
End Sub
Private Sub TCPSocket_Close()

 ' Cuando se Desconecta el Socket, el usuario para a estado no Logueado
 SocketTCP.CambiarEstadoDelCliente (0) ' Poner el Estado No Logueado
 
End Sub
Private Sub TCPSocket_DataArrival(ByVal BytesTotal As Long)
Dim DatosRecibidos

 ' **************************************************************
 ' Toma los datos que estan llegando al Socket
 ' **************************************************************
 TCPSocket.GetData DatosRecibidos, vbString, BytesTotal
 TratarPaquete DatosRecibidos, BytesTotal
 
End Sub
Private Sub TratarPaquete(DatosRecibidos As Variant, BytesTotal As Long)
Dim ComandoAccion, ComandoDatos, ResultadoComando As String
Dim PaqueteRecibido As Variant
Dim LargoDelPaquete As Integer

 ' **************************************************************
 ' Preprosesa el Paquete recibido a traves del Modulo (Recibir
 ' Paquete), en este punto se pueden hacer procesos de
 ' descompresion, etc.
 ' **************************************************************
 PaqueteRecibido = RecibirPaqueteTCP(DatosRecibidos)
 
 ' **************************************************************
 ' Todos los Paquetes como minimo deben medir 2 caracteres....
 ' Caracter 1:  Comando a Procesar
 ' Caracter 2:  Datos del Comando (Hasta el Final del Paquete)
 ' **************************************************************
 'Debug.Print PaqueteRecibido
 'Debug.Print PaqueteRecibido
 'Debug.Print PaqueteRecibido
 
 LargoDelPaquete = Len(PaqueteRecibido)
 ' Valida el Largo del Paquete
 If LargoDelPaquete < 2 Then
  ' ********
  ' El Paquete es Incorrecto se Descarta y no se procesa nada
  ' ********
  Exit Sub
 End If
 
 ' **************************************************************
 ' Separa el Paquete en la Accion y los Datos de la Accion
 ' **************************************************************
 ComandoAccion = Mid$(PaqueteRecibido, 1, 1)
 ComandoDatos = Mid$(PaqueteRecibido, 2)

 ' **************************************************************
 ' Ejecuta el Comando Solicitado
 ' **************************************************************
 Select Case ComandoAccion
  Case "0": ' Paquetes de Login
   ResultadoComando = ComandoAccion_0(ComandoDatos)
  Case "1": ' Paquetes de Estado
   ResultadoComando = ComandoAccion_1(ComandoDatos)
  Case "2": ' Paquetes de Listado de Usuario
   ResultadoComando = ComandoAccion_2(ComandoDatos)
  Case "3": ' Intercambio de Paquetes
   ResultadoComando = ComandoAccion_3(ComandoDatos)
  Case "4": ' Paquetes de Intercambio de Mensaje
   ResultadoComando = ComandoAccion_4(ComandoDatos)
 End Select

End Sub
Private Sub TCPSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 
  ' **************************************************************
  ' Si se produce un Error en el Socket Lo Cierra
  ' **************************************************************
  Cliente.TCPSocket.Close
  SocketTCP.CambiarEstadoDelCliente (0)
  
End Sub
Private Sub TiempoEnLinea_Click()
Dim Fecha, Hora As String

 ' **************************************************************
 ' Ejecuta el Sonido de Click
 ' **************************************************************
 Audio.EjecutarSonido "003"
 
 ' **************************************************************
 ' Si no esta logueado loinforma y sale...
 ' **************************************************************
 If Configuracion.Logueado <> 3 Then
  ' No se Encuentra Logueado al Sistema...
  MostrarMSGBox MensajeRecurso(197), vbOKOnly, "vbInformation", Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Muestra un Mensaje del Tiempo logueado
 ' **************************************************************
 Fecha = CompletarCadena(CStr(Day(TiempoEnLineaContanteDesde)), 2, "I", "0") & "/" & _
         CompletarCadena(CStr(Month(TiempoEnLineaContanteDesde)), 2, "I", "0") & "/" & _
         CompletarCadena(CStr(Year(TiempoEnLineaContanteDesde)), 4, "I", "0")
 Hora = CompletarCadena(CStr(Hour(TiempoEnLineaContanteDesde)), 2, "I", "0") & ":" & _
         CompletarCadena(CStr(Minute(TiempoEnLineaContanteDesde)), 2, "I", "0") & ":" & _
         CompletarCadena(CStr(Second(TiempoEnLineaContanteDesde)), 2, "I", "0")
 
 ' Logueado desde el % a las
 MostrarMSGBox MensajeRecurso(198) & Fecha & MensajeRecurso(199) & Hora, vbOKOnly, "vbInformation", Configuracion.TituloVentanas
 
End Sub
Private Sub SysIcon_NIError(ByVal ErrorNumber As Long)
  
 ' **************************************************************
 ' Evento de Error
 ' **************************************************************
 ' Hubo un Error Al Intentar Inicializar/Descargar el TrayIcon...
 MostrarMSGBox MensajeRecurso(200), vbOKOnly, "vbCritical", Configuracion.TituloVentanas

End Sub
Sub CargarLosMenus()

 ' **************************************************************
 ' Menu De Cambio de Grupo
 ' **************************************************************
  Set MenuDeCambioDeGrupo = New IcoMenu
  With MenuDeCambioDeGrupo
   ' Ver Mas...
   .SetItem 0, MensajeRecurso(188), ImagenesMenus.ListImages("MoverAGrupo").Picture
   .SetItem 1, ""
   ' Fuera de Grupo...
   .SetItem 2, MensajeRecurso(132), ImagenesMenus.ListImages("MoverAGrupo").Picture, ""
   ' **************************************************************
   ' Carga los Grupos
   ' **************************************************************
   Dim Contador, Cantidad As Integer
   Cantidad = 0
   For Contador = 1 To Cliente.ListadoDeAmigos.Nodes.Count
    If Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 1, 1) = "G" Then
     Cantidad = Cantidad + 1
     .SetItem 2 + Cantidad, Mid$(Cliente.ListadoDeAmigos.Nodes(Contador).key, 2), ImagenesMenus.ListImages("MoverAGrupo").Picture, ""
    End If
   Next
  End With
 
 ' **************************************************************
 ' Menu Estados del Usuario Menu Descolgable
 ' **************************************************************
 Set MenuCambioDeEstado = New IcoMenu
  With MenuCambioDeEstado
   ' Disponible (Normal)...
   .SetItem 0, MensajeRecurso(180), Imagenes.ListImages("EstadoVisible").Picture
   ' No Disponible...
   .SetItem 1, MensajeRecurso(181), Imagenes.ListImages("EstadoNoDisponible").Picture
   ' Enseguida Vuelvo...
   .SetItem 2, MensajeRecurso(182), Imagenes.ListImages("EstadoCustom").Picture
   ' No Molestar...
   .SetItem 3, MensajeRecurso(183), Imagenes.ListImages("EstadoCustom").Picture
   ' Definir Estado... (Custom)
   .SetItem 4, MensajeRecurso(184), Imagenes.ListImages("EstadoCustom").Picture
  End With
 
 ' **************************************************************
 ' Menu Estados del Usuario
 ' **************************************************************
 Set MenuDeEstadosDeUsuario = New IcoMenu
  With MenuDeEstadosDeUsuario
   ' Disponible (Normal)...
   .SetItem 0, MensajeRecurso(180), Imagenes.ListImages("EstadoVisible").Picture
   ' No Disponible...
   .SetItem 1, MensajeRecurso(181), Imagenes.ListImages("EstadoNoDisponible").Picture
   ' Enseguida Vuelvo...
   .SetItem 2, MensajeRecurso(182), Imagenes.ListImages("EstadoCustom").Picture
   ' No Molestar...
   .SetItem 3, MensajeRecurso(183), Imagenes.ListImages("EstadoCustom").Picture
   ' Definir Estado... (Custom)
   .SetItem 4, MensajeRecurso(184), Imagenes.ListImages("EstadoCustom").Picture
  End With

 ' **************************************************************
 ' Menu Coneccion
 ' **************************************************************
 Set MenuToolConeccion = New IcoMenu
 With MenuToolConeccion
   ' Conectar
   .SetItem 0, MensajeRecurso(107), ImagenesMenus.ListImages("Conectar").Picture
   ' Desconectar
   .SetItem 1, MensajeRecurso(202), ImagenesMenus.ListImages("Desconectar").Picture
   ' Cambiar Usuario
   .SetItem 2, MensajeRecurso(203), ImagenesMenus.ListImages("CambiarUsuario").Picture
   .SetItem 3, ""
   ' Salir
   .SetItem 4, MensajeRecurso(204), ImagenesMenus.ListImages("Salir").Picture
 End With

 ' **************************************************************
 ' Menu Configuracion
 ' **************************************************************
 Set MenuToolConfiguracion = New IcoMenu
 With MenuToolConfiguracion
   ' Ver Mis Datos...
   .SetItem 0, MensajeRecurso(205), ImagenesMenus.ListImages("MostrarMisDatos").Picture
   ' Cambiar Mis Datos...
   .SetItem 1, MensajeRecurso(206), ImagenesMenus.ListImages("CambiarMisDatos").Picture
   .SetItem 2, ""
   ' Cambiar Password...
   .SetItem 3, MensajeRecurso(154) & "...", ImagenesMenus.ListImages("Clave").Picture
   .SetItem 4, ""
   ' Preferencias...
   .SetItem 5, MensajeRecurso(208), ImagenesMenus.ListImages("Preferencias").Picture
 End With

 ' **************************************************************
 ' Menu Amigos
 ' **************************************************************
 Set MenuToolAmigos = New IcoMenu
 With MenuToolAmigos
   ' Recargar Amigos...
   .SetItem 0, MensajeRecurso(209), ImagenesMenus.ListImages("RecargarAmigos").Picture
   .SetItem 1, ""
   ' Crear Grupo...
   .SetItem 2, MensajeRecurso(167), ImagenesMenus.ListImages("GrupoAgregar").Picture
   ' Eliminar Grupo...
   .SetItem 3, MensajeRecurso(211), ImagenesMenus.ListImages("GrupoEliminar").Picture
   .SetItem 4, ""
   ' Agregar Amigo...
   .SetItem 5, MensajeRecurso(130), ImagenesMenus.ListImages("AmigoAgregar").Picture
   ' Eliminar Amigo...
   .SetItem 6, MensajeRecurso(213), ImagenesMenus.ListImages("AmigoEliminar").Picture
   ' Buscar Amigos...
   .SetItem 7, MensajeRecurso(169), ImagenesMenus.ListImages("BuscarAmigos").Picture
   .SetItem 8, ""
   ' Bloqueo de Amigos...
   .SetItem 9, MensajeRecurso(215), ImagenesMenus.ListImages("BloqueoDeAmigos").Picture
 End With

 ' **************************************************************
 ' Menu Ayuda
 ' **************************************************************
 Set MenuToolAyuda = New IcoMenu
 With MenuToolAyuda
   ' Ayuda en Linea...
   .SetItem 0, MensajeRecurso(216), ImagenesMenus.ListImages("Ayuda").Picture
   .SetItem 1, ""
   ' Sobre EIM...
   .SetItem 2, MensajeRecurso(217), ImagenesMenus.ListImages("EIM").Picture
 End With

 ' **************************************************************
 ' Menu ClickAmigos
 ' **************************************************************
 Set MenuClickAmigo = New IcoMenu
 With MenuClickAmigo
   ' Enviar Mensaje...
   .SetItem 0, MensajeRecurso(218), ImagenesMenus.ListImages("EIM").Picture
   ' Enviar Archivo...
   '.SetItem 1, MensajeRecurso(219), ImagenesMenus.ListImages("EnviarArchivo").Picture
   .SetItem 1, ""
   ' Ver Datos...
   .SetItem 2, MensajeRecurso(220), ImagenesMenus.ListImages("MostrarMisDatos").Picture
   .SetItem 3, ""
   ' Eliminar Amigo...
   .SetItem 4, MensajeRecurso(213), ImagenesMenus.ListImages("AmigoEliminar").Picture
   .SetItem 5, ""
   ' Mover a Grupo...
   .SetItem 6, MensajeRecurso(147), ImagenesMenus.ListImages("MoverAGrupo").Picture, , True
   .SetItem 7, ""
   ' Enviar Mail...
   .SetItem 8, MensajeRecurso(452), ImagenesMenus.ListImages("EnviarMail").Picture
 End With

 ' **************************************************************
 ' Menu ClickGrupo
 ' **************************************************************
 Set MenuClickGrupo = New IcoMenu
 With MenuClickGrupo
   ' Enviar Mensaje Grupal...
   '.SetItem 0, MensajeRecurso(222), ImagenesMenus.ListImages("MensajeGrupal").Picture
   '.SetItem 1, ""
   ' Eliminar Grupo...
   .SetItem 0, MensajeRecurso(211), ImagenesMenus.ListImages("GrupoEliminar").Picture
   ' Renombrar Grupo...
   .SetItem 1, MensajeRecurso(224), ImagenesMenus.ListImages("RenombrarGrupo").Picture
   .SetItem 2, ""
   ' Agregar Amigo...
   .SetItem 3, MensajeRecurso(130), ImagenesMenus.ListImages("AmigoAgregar").Picture
 End With

 ' **************************************************************
 ' Menu ClickUsuario
 ' **************************************************************
 Set MenuClickUsuario = New IcoMenu
 With MenuClickUsuario
   ' Cambiar Mis Datos...
   .SetItem 0, MensajeRecurso(206), ImagenesMenus.ListImages("CambiarMisDatos").Picture
   ' Ver Mis Datos...
   .SetItem 1, MensajeRecurso(205), ImagenesMenus.ListImages("MostrarMisDatos").Picture
   ' Recargar Amigos...
   .SetItem 2, MensajeRecurso(209), ImagenesMenus.ListImages("RecargarAmigos").Picture
   .SetItem 3, ""
   ' Crear Grupo...
   .SetItem 4, MensajeRecurso(167), ImagenesMenus.ListImages("GrupoAgregar").Picture
   ' Eliminar Grupo...
   .SetItem 5, MensajeRecurso(211), ImagenesMenus.ListImages("GrupoEliminar").Picture
   .SetItem 6, ""
   ' Agregar Amigo...
   .SetItem 7, MensajeRecurso(130), ImagenesMenus.ListImages("AmigoAgregar").Picture
   ' Eliminar Amigo...
   .SetItem 8, MensajeRecurso(213), ImagenesMenus.ListImages("AmigoEliminar").Picture
   ' Eliminar Amigo...
   .SetItem 9, MensajeRecurso(169), ImagenesMenus.ListImages("BuscarAmigos").Picture
   ' Bloqueo de Amigos...
   .SetItem 10, MensajeRecurso(215), ImagenesMenus.ListImages("BloqueoDeAmigos").Picture
   .SetItem 11, ""
   ' Preferencias...
   .SetItem 12, MensajeRecurso(208), ImagenesMenus.ListImages("Preferencias").Picture
 End With

 ' **************************************************************
 ' Menu ClickUsuario (Especial Para Cuando esta Desconectador o
 ' Conectando...)
 ' **************************************************************
 Set MenuClickUsuarioDesconectadoConectando = New IcoMenu
 With MenuClickUsuarioDesconectadoConectando
   ' Conectar...
   .SetItem 0, MensajeRecurso(107), ImagenesMenus.ListImages("Conectar").Picture
   ' Preferencias...
   .SetItem 1, MensajeRecurso(208), ImagenesMenus.ListImages("Preferencias").Picture
   ' Salir...
   .SetItem 2, MensajeRecurso(204), ImagenesMenus.ListImages("Salir").Picture
 End With

 ' **************************************************************
 ' Menu PopUpTary
 ' **************************************************************
 Set MenuPopUpTray = New IcoMenu
 With MenuPopUpTray
   ' Conectar...
   .SetItem 0, MensajeRecurso(107), ImagenesMenus.ListImages("Conectar").Picture
   ' Desconectar...
   .SetItem 1, MensajeRecurso(202), ImagenesMenus.ListImages("Desconectar").Picture
   ' Cambiar Usuario...
   .SetItem 2, MensajeRecurso(203), ImagenesMenus.ListImages("CambiarUsuario").Picture
   ' Preferencias...
   .SetItem 3, MensajeRecurso(208), ImagenesMenus.ListImages("Preferencias").Picture
   .SetItem 4, ""
   ' Mostrar EIM...
   .SetItem 5, MensajeRecurso(226), ImagenesMenus.ListImages("MensajeGrupal").Picture
   ' Sobre EIM...
   .SetItem 6, MensajeRecurso(217), ImagenesMenus.ListImages("EIM").Picture
   .SetItem 7, ""
   ' Cambiar Estado...
   .SetItem 8, MensajeRecurso(227) & "  ", ImagenesMenus.ListImages("CambiarEstado").Picture, , True
   .SetItem 9, ""
   ' Salir...
   .SetItem 10, MensajeRecurso(204), ImagenesMenus.ListImages("Salir").Picture
 End With


End Sub
Private Sub TimerMensaje_Timer()
Dim Texto As String

 ' **************************************************************
 ' Decide si Esta Conectado, Conectando o Desconectado...
 ' **************************************************************
 Select Case Configuracion.Logueado
  Case 0:
   ' DesConectado
   Texto = MensajeRecursoReal(228)  ' Esto es para que traiga Desconectado en
                                    ' español...
  Case 1:
   ' Conectando
   Texto = MensajeRecursoReal(229)  ' Esto es para que traiga Desconectado en
                                    ' español...
  Case 3:
   ' Conectado
   Texto = MensajeRecursoReal(230)  ' Esto es para que traiga Desconectado en
                                    ' español...
 End Select
 
 ' **************************************************************
 ' Flashea el Mensaje
 ' **************************************************************
 If TimerMensaje Then
  ' **************************************************************
  ' Verifica que existe el Nodo que debe Cambiarse
  ' **************************************************************
  If Mid$(Cliente.ListadoDeAmigos.Nodes(1).Text, 1, 1) <> "*" Then
   Exit Sub
  End If
  Select Case Me.IndiceTimerMensaje
   Case 0
    Cliente.ListadoDeAmigos.Nodes(1).Image = "Mensaje"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "Mensaje"
    Varios.CambiarIconoTray Texto & "ConMensaje"
   Case 1
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash"
   Case 2
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash2"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash2"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash2"
   Case 3
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash3"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash3"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash3"
   Case 4
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash3"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash3"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash3"
   Case 5
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash4"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash4"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash4"
   Case 6
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash5"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash5"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash5"
   Case 7
    Cliente.ListadoDeAmigos.Nodes(1).Image = "MensajeFlash6"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "MensajeFlash6"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash6"
   Case 8
    Cliente.ListadoDeAmigos.Nodes(1).Image = "Mensaje0"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "Mensaje0"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash0"
   Case 9
    Cliente.ListadoDeAmigos.Nodes(1).Image = "Mensaje0"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "Mensaje0"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash0"
   Case 10
    Cliente.ListadoDeAmigos.Nodes(1).Image = "Mensaje0"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "Mensaje0"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash0"
   Case 11
    Cliente.ListadoDeAmigos.Nodes(1).Image = "Mensaje0"
    Cliente.ListadoDeAmigos.Nodes(1).SelectedImage = "Mensaje0"
    Varios.CambiarIconoTray Texto & "ConMensajeFlash0"
  End Select
  IndiceTimerMensaje = IndiceTimerMensaje + 1
  If IndiceTimerMensaje > 11 Then IndiceTimerMensaje = 0
 End If

End Sub
