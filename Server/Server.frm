VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form Server 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   32
      Top             =   5505
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock TCPSocket 
      Index           =   0
      Left            =   2700
      Top             =   4470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab Lenguetas 
      Height          =   5505
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9710
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Monitor"
      TabPicture(0)   =   "Server.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Configuration"
      TabPicture(1)   =   "Server.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "User's"
      TabPicture(2)   =   "Server.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "RefrescoUsuarioActuales"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).Control(3)=   "TiempoDeRefresco"
      Tab(2).Control(4)=   "Frame9"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "User Manager"
      TabPicture(3)   =   "Server.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(2)=   "Frame12"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame12 
         Height          =   675
         Left            =   -74880
         TabIndex        =   60
         Top             =   4740
         Width           =   9405
         Begin VB.CommandButton UserManCrearUsuario 
            Caption         =   "Add User..."
            Height          =   435
            Left            =   7530
            TabIndex        =   19
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton UserManBorrarUsuario 
            Caption         =   "Delete User..."
            Height          =   435
            Left            =   5580
            TabIndex        =   18
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton UserManGrabarCambios 
            Caption         =   "Save Profile..."
            Height          =   435
            Left            =   3660
            TabIndex        =   17
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton UserManDescartarCambios 
            Caption         =   "Discard Change's..."
            Height          =   435
            Left            =   1710
            TabIndex        =   16
            Top             =   180
            Width           =   1785
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Datos Del Usuario "
         Height          =   4290
         Left            =   -72030
         TabIndex        =   59
         Top             =   480
         Width           =   6555
         Begin NTService.NTService NTService 
            Left            =   270
            Top             =   3330
            _Version        =   65536
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   0
            ServiceName     =   "Simple"
            StartMode       =   3
         End
         Begin VB.CommandButton Command1 
            Caption         =   "->"
            Height          =   345
            Left            =   5880
            TabIndex        =   79
            Top             =   3930
            Width           =   345
         End
         Begin VB.ComboBox User_EstadoCivil 
            Height          =   315
            ItemData        =   "Server.frx":093A
            Left            =   1560
            List            =   "Server.frx":094A
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1950
            Width           =   1725
         End
         Begin VB.ComboBox User_Sexo 
            Height          =   315
            ItemData        =   "Server.frx":0972
            Left            =   1560
            List            =   "Server.frx":097C
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3600
            Width           =   1275
         End
         Begin VB.TextBox User_Password 
            Height          =   315
            Left            =   3600
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "User_Password"
            Top             =   3930
            Width           =   2235
         End
         Begin VB.ComboBox User_Signo 
            Height          =   315
            ItemData        =   "Server.frx":098E
            Left            =   4260
            List            =   "Server.frx":09B6
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1620
            Width           =   2175
         End
         Begin VB.CheckBox User_UsuarioBloqueado 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   3960
            Width           =   225
         End
         Begin VB.TextBox User_UbicacionGeografica 
            Height          =   315
            Left            =   4890
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "User_UbicacionGeografica"
            Top             =   1950
            Width           =   1545
         End
         Begin VB.TextBox User_Telefono 
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   13
            Text            =   "User_Telefono"
            Top             =   3600
            Width           =   2865
         End
         Begin VB.TextBox User_OtraInfo 
            Height          =   315
            Left            =   1560
            MaxLength       =   150
            TabIndex        =   11
            Text            =   "User_OtraInfo"
            Top             =   3270
            Width           =   4905
         End
         Begin VB.TextBox User_Ocupacion 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "User_Ocupacion"
            Top             =   2940
            Width           =   2655
         End
         Begin VB.TextBox User_Intencion 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   9
            Text            =   "User_Intencion"
            Top             =   2610
            Width           =   2655
         End
         Begin VB.TextBox User_Humor 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   8
            Text            =   "User_Humor"
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox User_Edad 
            Height          =   315
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   4
            Text            =   "User_Edad"
            Top             =   1620
            Width           =   825
         End
         Begin VB.TextBox User_ApellidoYNombre 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "User_ApellidoYNombre"
            Top             =   630
            Width           =   4575
         End
         Begin VB.TextBox User_DireccionDeEmail 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "User_DireccionDeEmail"
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox User_FechadeNacimiento 
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "User_FechadeNacimiento"
            Top             =   1290
            Width           =   1455
         End
         Begin VB.TextBox User_IDAliasUsuario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   0
            Text            =   "User_IDAliasUsuario"
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label25 
            Caption         =   "Geografical Site:"
            Height          =   285
            Left            =   3450
            TabIndex        =   78
            Top             =   1980
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Signo:"
            Height          =   285
            Left            =   3420
            TabIndex        =   77
            Top             =   1650
            Width           =   1485
         End
         Begin VB.Label Label23 
            Caption         =   "Password:"
            Height          =   285
            Left            =   2790
            TabIndex        =   76
            Top             =   3960
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Phone:"
            Height          =   285
            Left            =   3030
            TabIndex        =   75
            Top             =   3630
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Block User:"
            Height          =   285
            Left            =   90
            TabIndex        =   74
            Top             =   3960
            Width           =   1485
         End
         Begin VB.Label Label20 
            Caption         =   "Sex:"
            Height          =   285
            Left            =   90
            TabIndex        =   73
            Top             =   3630
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Other Info:"
            Height          =   285
            Left            =   90
            TabIndex        =   72
            Top             =   3300
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Ocupation:"
            Height          =   285
            Left            =   90
            TabIndex        =   71
            Top             =   2970
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "Intention:"
            Height          =   285
            Left            =   90
            TabIndex        =   70
            Top             =   2640
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "User Humor:"
            Height          =   285
            Left            =   90
            TabIndex        =   69
            Top             =   2310
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Civil State:"
            Height          =   285
            Left            =   90
            TabIndex        =   68
            Top             =   1980
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Age:"
            Height          =   285
            Left            =   90
            TabIndex        =   67
            Top             =   1650
            Width           =   1485
         End
         Begin VB.Label Label13 
            Caption         =   "BirthDay:"
            Height          =   285
            Left            =   90
            TabIndex        =   66
            Top             =   1320
            Width           =   1485
         End
         Begin VB.Label Label12 
            Caption         =   "E-Mail Direction:"
            Height          =   285
            Left            =   90
            TabIndex        =   65
            Top             =   990
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "Name:"
            Height          =   285
            Left            =   90
            TabIndex        =   64
            Top             =   660
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "User Alias: "
            Height          =   285
            Left            =   90
            TabIndex        =   63
            Top             =   330
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Usuarios Registrado"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   2805
         Begin VB.ListBox UsuariosRegistrados 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   3735
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   2610
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   3765
            Left            =   90
            Top             =   240
            Width           =   2625
         End
      End
      Begin VB.Frame Frame9 
         Height          =   675
         Left            =   -72630
         TabIndex        =   52
         Top             =   4740
         Width           =   7155
         Begin VB.CommandButton DesconectarUsuario 
            Caption         =   "Disconnect User..."
            Height          =   435
            Left            =   3360
            TabIndex        =   54
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton RefrescarListadoDeUsuarios 
            Caption         =   "Refresh Now..."
            Height          =   435
            Left            =   5280
            TabIndex        =   53
            Top             =   180
            Width           =   1785
         End
      End
      Begin VB.TextBox TiempoDeRefresco 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -74790
         TabIndex        =   47
         Text            =   "5000"
         Top             =   4980
         Width           =   735
      End
      Begin VB.Frame Frame8 
         Caption         =   "Refresh Time..."
         Height          =   675
         Left            =   -74880
         TabIndex        =   50
         Top             =   4740
         Width           =   2205
         Begin VB.Label Label5 
            Caption         =   "Milliseconds..."
            Height          =   285
            Left            =   870
            TabIndex        =   51
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Timer RefrescoUsuarioActuales 
         Interval        =   5000
         Left            =   -70170
         Top             =   3540
      End
      Begin VB.Frame Frame7 
         Caption         =   "Actual Logued User's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   9405
         Begin VB.ListBox ListadoUsuariosActuales 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   3735
            ItemData        =   "Server.frx":0A1F
            Left            =   90
            List            =   "Server.frx":0A26
            Sorted          =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   9210
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   3765
            Left            =   90
            Top             =   240
            Width           =   9225
         End
      End
      Begin VB.Frame Frame6 
         Height          =   675
         Left            =   -74880
         TabIndex        =   43
         Top             =   4740
         Width           =   9405
         Begin VB.CommandButton DescartarCambios 
            Caption         =   "Discard Changes..."
            Height          =   435
            Left            =   3660
            TabIndex        =   46
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton GrabarCambios 
            Caption         =   "Save..."
            Height          =   435
            Left            =   5610
            TabIndex        =   45
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton ReiniciarSistema 
            Caption         =   "Restart System..."
            Height          =   435
            Left            =   7530
            TabIndex        =   44
            Top             =   180
            Width           =   1785
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Actual Configuration"
         Height          =   4245
         Left            =   -74910
         TabIndex        =   38
         Top             =   420
         Width           =   9435
         Begin VB.TextBox DireccionEMAILAdministrador 
            Height          =   315
            Left            =   2940
            MaxLength       =   100
            TabIndex        =   25
            Text            =   "DireccionEMAILAdministrador"
            Top             =   2730
            Width           =   5025
         End
         Begin VB.TextBox DireccionSMTP 
            Height          =   315
            Left            =   2940
            MaxLength       =   15
            TabIndex        =   24
            Text            =   "DireccionSMTP"
            Top             =   2280
            Width           =   1905
         End
         Begin VB.CheckBox PermitirCrearUsuarios 
            Height          =   315
            Left            =   2940
            TabIndex        =   26
            Top             =   3210
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox UsuariosSoportados 
            Height          =   315
            Left            =   2940
            MaxLength       =   3
            TabIndex        =   22
            Text            =   "UsuariosSoportados "
            Top             =   1380
            Width           =   1905
         End
         Begin VB.TextBox NombreDeLaBaseDeDatos 
            Height          =   315
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   21
            Text            =   "NombreDeLaBaseDeDatos"
            Top             =   930
            Width           =   3645
         End
         Begin VB.TextBox PortTCP 
            Height          =   315
            Left            =   2940
            MaxLength       =   5
            TabIndex        =   23
            Text            =   "PortTCP"
            Top             =   1830
            Width           =   1905
         End
         Begin VB.TextBox UbicacionBaseDeDatos 
            Height          =   315
            Left            =   2940
            MaxLength       =   200
            TabIndex        =   20
            Text            =   "UbicacionBaseDeDatos"
            Top             =   480
            Width           =   5025
         End
         Begin VB.Label Label9 
            Caption         =   "(Option Not Implemented Yet...)"
            Height          =   240
            Left            =   3180
            TabIndex        =   62
            Top             =   3250
            Width           =   2430
         End
         Begin VB.Label Label6 
            Caption         =   "Administrator E-Mail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Top             =   2790
            Width           =   2685
         End
         Begin VB.Label Label8 
            Caption         =   "The Client Can Create Users:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   56
            Top             =   3270
            Width           =   2685
         End
         Begin VB.Label Label7 
            Caption         =   "SMTP Server (IP Direction):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   55
            Top             =   2334
            Width           =   2685
         End
         Begin VB.Label Label4 
            Caption         =   "TCP Port:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   42
            Top             =   1878
            Width           =   2685
         End
         Begin VB.Label Label3 
            Caption         =   "Maximus User's:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   1422
            Width           =   2685
         End
         Begin VB.Label Label2 
            Caption         =   "Database Name (MDB File):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   966
            Width           =   2685
         End
         Begin VB.Label Label1 
            Caption         =   "Database Path:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   510
            Width           =   2685
         End
      End
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   4650
         TabIndex        =   34
         Top             =   4740
         Width           =   4875
         Begin VB.CommandButton CerrarSistema 
            Caption         =   "Exit..."
            Height          =   435
            Left            =   120
            TabIndex        =   37
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton BotonArranque 
            Caption         =   "Stop System..."
            Height          =   435
            Left            =   3000
            TabIndex        =   35
            Top             =   180
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Current User's"
         Height          =   675
         Left            =   2460
         TabIndex        =   31
         Top             =   4740
         Width           =   2055
         Begin VB.Image Image1 
            Height          =   240
            Left            =   120
            Picture         =   "Server.frx":0A43
            Stretch         =   -1  'True
            Top             =   270
            Width           =   240
         End
         Begin VB.Label CantidadUsuariosLabel 
            Caption         =   "Usuarios x/x"
            Height          =   225
            Left            =   450
            TabIndex        =   36
            Top             =   330
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "System State"
         Height          =   675
         Left            =   120
         TabIndex        =   30
         Top             =   4740
         Width           =   2205
         Begin VB.Image EstadoImagen 
            Height          =   240
            Left            =   90
            Picture         =   "Server.frx":0FCD
            Top             =   270
            Width           =   240
         End
         Begin VB.Label EstadoLabel 
            Height          =   285
            Left            =   420
            TabIndex        =   33
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "System Log's"
         Height          =   4245
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   9405
         Begin MSWinsockLib.Winsock SMTPMailSocket 
            Left            =   3660
            Top             =   3060
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList EstadoSistema 
            Left            =   1590
            Top             =   3240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   12
            ImageHeight     =   12
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Server.frx":1557
                  Key             =   "Levantado"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Server.frx":1AF1
                  Key             =   "Detenido"
               EndProperty
            EndProperty
         End
         Begin RichTextLib.RichTextBox MensajesServidor 
            Height          =   3880
            Left            =   110
            TabIndex        =   29
            Top             =   250
            Width           =   9160
            _ExtentX        =   16166
            _ExtentY        =   6826
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            DisableNoScroll =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Server.frx":208B
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
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   3945
            Left            =   90
            Top             =   240
            Width           =   9225
         End
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IndiceActual As Integer
Private Sub NTService_Start(Success As Boolean)
On Error GoTo ErrorServicio
    
    Success = True
    Exit Sub

ErrorServicio:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub
Private Sub BotonArranque_Click()
Dim Respuesta As Integer

 ' **************************************************************
 ' Segun el Estado del Sistema lo Sube o Baja...
 ' **************************************************************
 If Configuracion.EstadoDelSistema = "Up" Then
   Respuesta = MsgBox("¿Are you Sure to Down The System?", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
   ' Pregunta si se esta seguro de la bajada del sistema, si contesta "No"
   ' cancela la bajada...
   If Respuesta = vbNo Then
    Exit Sub
   End If
   DetenerSistema
  Else
   InicializarSistema
 End If
 
End Sub
Private Sub NTService_Stop()
On Error GoTo ErrorServicio
    
 ' ************************************************************************
 ' Bajar Todo...
 ' ************************************************************************
 DetenerSistema
 BajarBaseDeDatos
 Unload Me

ErrorServicio:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub
Private Sub CerrarSistema_Click()
Dim Respuesta As Integer

 Respuesta = MsgBox("¿Are you Sure to Close The System?", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 ' Pregunta si se esta seguro de la bajada del sistema, si contesta "No"
 ' cancela la bajada...
 If Respuesta = vbNo Then
  Exit Sub
 End If
 
 NTService.StopService
   
End Sub
Private Sub Command1_Click()
Dim Respuesta As Boolean

 ' **************************************************************
 ' Envia un EMail con Usuario y Password...
 ' **************************************************************
 Dim Respuestas As Boolean
 Respuesta = Varios.EnviarPasswordAUsuario(Trim(Me.User_IDAliasUsuario), Trim(Me.User_Password), Trim(Me.User_DireccionDeEmail), Trim(Me.User_ApellidoYNombre))
 If Respuesta = False Then
  MsgBox "An Error Ocurred When Try to Send the Password to [" & Trim(Me.User_IDAliasUsuario) & "] to [" & Trim(Me.User_DireccionDeEmail) & "]...", vbOKOnly + vbCritical, Configuracion.TituloVentanas
 End If

End Sub
Private Sub DescartarCambios_Click()
Dim Respuesta As Integer

 ' **************************************************************
 ' Verifica si hubo Cambios...
 ' **************************************************************
 Respuesta = 0
 If UCase(Trim(Server.UbicacionBaseDeDatos)) <> UCase(Trim(Configuracion.UbicacionBaseDeDatos)) Then Respuesta = 1
 If UCase(Trim(Server.NombreDeLaBaseDeDatos)) <> UCase(Trim(Configuracion.NombreDeLaBaseDeDatos)) Then Respuesta = 1
 If UCase(Trim(Server.UsuariosSoportados)) <> UCase(Trim(Configuracion.UsuariosSoportados)) Then Respuesta = 1
 If UCase(Trim(Server.PortTCP)) <> UCase(Trim(Configuracion.PortTCP)) Then Respuesta = 1
 If UCase(Trim(Server.DireccionSMTP)) <> UCase(Trim(Configuracion.DireccionIPSMTP)) Then Respuesta = 1
 If UCase(Trim(Server.DireccionEMAILAdministrador)) <> UCase(Trim(Configuracion.DireccionEMAILAdministrador)) Then Respuesta = 1
 If UCase(Trim(Server.PermitirCrearUsuarios)) <> UCase(Trim(Configuracion.PermitirCrear)) Then Respuesta = 1
 If Respuesta = 0 Then Exit Sub ' No hay Cambios que descartar...
 
 ' **************************************************************
 ' Descarta los Cambios Realizados en la Configuracion...
 ' **************************************************************
 ' Verifica que este seguro de deshacer los cambios...
 Respuesta = MsgBox("¿Do you Like to Discard the Change's?", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 If Respuesta = vbNo Then
  Exit Sub
 End If
  
 ' **************************************************************
 ' Deshace los Cambios
 ' **************************************************************
 Server.UbicacionBaseDeDatos = Configuracion.UbicacionBaseDeDatos
 Server.NombreDeLaBaseDeDatos = Configuracion.NombreDeLaBaseDeDatos
 Server.UsuariosSoportados = Configuracion.UsuariosSoportados
 Server.PortTCP = Configuracion.PortTCP
 Server.DireccionSMTP = Configuracion.DireccionIPSMTP
 Server.DireccionEMAILAdministrador = Configuracion.DireccionEMAILAdministrador
 Server.PermitirCrearUsuarios = Configuracion.PermitirCrear
  
End Sub
Private Sub DesconectarUsuario_Click()
Dim Respuesta As Integer
Dim Posicion, Usuario As String

 ' ************************************************************************
 ' Verifica que el Sistema Este Levantado...
 ' ************************************************************************
 If Configuracion.EstadoDelSistema <> "Up" Then
  MsgBox "The System is Not Running...", vbCritical + vbOKOnly, Configuracion.TituloVentanas
  Exit Sub
 End If
 ' ************************************************************************
 
 ' ************************************************************************
 ' Levanta el Texto de la Lista (Usuario, etc.)
 ' ************************************************************************
 Usuario = Server.ListadoUsuariosActuales.List(Server.ListadoUsuariosActuales.ListIndex)
 ' ************************************************************************
 
 ' ************************************************************************
 ' Verifica que Sea un Usuario
 ' ************************************************************************
 If Len(Usuario) < 40 Then
  ' Se produjo un Error en la Seleccion...
  MsgBox "Please Select one User to Disconnect...", vbCritical + vbOKOnly, Configuracion.TituloVentanas
  Exit Sub
 End If
 ' Aca se verifica que sea un usuario, ya que cuando se carga la lista de
 ' usuarios, a cada usuario logueado se le agrega "- "
 If Mid$(Usuario, 1, 2) = "-" Then
  ' Se produjo un Error en la Seleccion...
  MsgBox "Please Select one User to Disconnect...", vbCritical + vbOKOnly, Configuracion.TituloVentanas
  Exit Sub
 End If
 ' ************************************************************************
 
 ' ************************************************************************
 ' Verifica que este seguro de queres echarlo...
 ' ************************************************************************
 Posicion = Mid$(Usuario, 32, 5)
 Usuario = Trim(Mid$(Usuario, 3, 16))
 Respuesta = MsgBox("¿Are you Sure to Disconnect the User [" & Usuario & "]?", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 ' Se arrepintio
 If Respuesta = vbNo Then
  Exit Sub
 End If
 
 ' ************************************************************************
 ' Busca en que port Esta y lo cierra...
 ' El tema de cambiar el estado del Port, etc, etc, no es preocupante, ya
 ' que de esto se encarga el evento "TCPSocket_Close"
 ' ************************************************************************
 ' Cierra el Socket...
 TCPSocket(CInt(Posicion)).Close
 TCPSocket_Close (CInt(Posicion))
 CargarListadoUsuariosActuales
 
End Sub
Private Sub Form_Load()
On Error GoTo ErrorServicio
Dim NombreDelServicio As String

 ' ************************************************************************
 ' Carga la Configuracion del Sistema Automaticamente...
 ' ************************************************************************
 Inicializacion.InicializarSistema
 
 ' ************************************************
 ' Levantar minimizado...
 ' ************************************************
 Me.WindowState = 1
 
 ' ************************************************************************
 ' Levanta el Servicio...
 ' ************************************************************************
 NTService.StartMode = svcStartAutomatic
 NTService.DisplayName = "EIM Version " & Configuracion.VersionDelSistema
 NombreDelServicio = NTService.DisplayName
 ' Si se pone de opcion /I instala el servicio
 If Mid$(UCase$(Trim$(Command$)), 1, 2) = UCase("/I") Then
  NTService.Interactive = True
  If NTService.Install Then
    NTService.SaveSetting "Parameters", "Versión", Configuracion.VersionDelSistema
    MsgBox NombreDelServicio & " Installed...", vbOKOnly + vbInformation, Configuracion.TituloVentanas
   Else
    MsgBox NombreDelServicio & " Can't be Installed...", vbOKOnly + vbCritical, Configuracion.TituloVentanas
  End If
  End
 End If
 ' Si se pone de opcion /D Desinstala el servicio
 If UCase$(Trim$(Command$)) = UCase("/D") Then
   If NTService.Uninstall Then
     MsgBox NombreDelServicio & " De-Installed...", vbOKOnly + vbInformation, Configuracion.TituloVentanas
     End
    Else
     MsgBox NombreDelServicio & " Can't be De-Installed...", vbOKOnly + vbCritical, Configuracion.TituloVentanas
     End
    End If
 End If
 ' Si pone algo Mal.... Sale sin Hacer Nada...
 If UCase$(Command$) <> "" Then
  MsgBox "Command Line Incorrect...", vbOKOnly + vbCritical, "NTSincro (Server)"
  End
 End If
 ' Startea el Servicio...
 NTService.StartService ' Confirma que el Servicio Levanto...

 ' ************************************************************************
 ' Carga la Pantalla de Configuracion
 ' ************************************************************************
 CargarPantallasAutomaticas
 
 ' ************************************************************************
 ' Pone el Nombre del Formulario...
 ' ************************************************************************
 Me.Caption = Configuracion.TituloVentanas
 
 ' ************************************************************************
 ' Se posiciona en la Lengueta 0 (Monitor)
 ' ************************************************************************
 Lenguetas.Tab = 0
 
 Exit Sub
 
ErrorServicio:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub
Private Sub SysIcon_NIError(ByVal ErrorNumber As Long)
  
  MsgBox "An Error Ocurred When Tar tu Load/Unload SysTray. Error=" & ErrorNumber

End Sub
Private Sub GrabarCambios_Click()
Dim Respuesta As Integer

 ' **************************************************************
 ' Verifica si hubo Cambios...
 ' **************************************************************
 Respuesta = 0
 If UCase(Trim(Server.UbicacionBaseDeDatos)) <> UCase(Trim(Configuracion.UbicacionBaseDeDatos)) Then Respuesta = 1
 If UCase(Trim(Server.NombreDeLaBaseDeDatos)) <> UCase(Trim(Configuracion.NombreDeLaBaseDeDatos)) Then Respuesta = 1
 If UCase(Trim(Server.UsuariosSoportados)) <> UCase(Trim(Configuracion.UsuariosSoportados)) Then Respuesta = 1
 If UCase(Trim(Server.PortTCP)) <> UCase(Trim(Configuracion.PortTCP)) Then Respuesta = 1
 If UCase(Trim(Server.DireccionSMTP)) <> UCase(Trim(Configuracion.DireccionIPSMTP)) Then Respuesta = 1
 If UCase(Trim(Server.DireccionEMAILAdministrador)) <> UCase(Trim(Configuracion.DireccionEMAILAdministrador)) Then Respuesta = 1
 If UCase(Trim(Server.PermitirCrearUsuarios)) <> UCase(Trim(Configuracion.PermitirCrear)) Then Respuesta = 1
 If Respuesta = 0 Then
  MsgBox "Are Not Change's to Save...", vbInformation, Configuracion.TituloVentanas
  Exit Sub ' No hay Cambios que descartar...
 End If
 
 ' **************************************************************
 ' Graba los Cambios Realizados
 ' **************************************************************
 ' Verifica que este seguro de Grabar y Restartear el Sistema
 Respuesta = MsgBox("¿Are you Sure to Save the Changes? (This Action Restart the System...)", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 If Respuesta = vbNo Then
  Exit Sub
 End If
  
 ' **************************************************************
 ' Si se ingresaron mas de 300 usuario concurrentes
 ' corrige el error
 ' **************************************************************
 If Configuracion.UsuariosSoportados > 300 Then
  Configuracion.UsuariosSoportados = 300
 End If
 
 ' **************************************************************
 ' Detiene el Sistema
 ' **************************************************************
 DetenerSistema
 
 ' **************************************************************
 ' Define los Nuevos parametros de Configuracion
 ' **************************************************************
 Configuracion.UbicacionBaseDeDatos = Server.UbicacionBaseDeDatos
 Configuracion.NombreDeLaBaseDeDatos = Server.NombreDeLaBaseDeDatos
 Configuracion.UsuariosSoportados = Server.UsuariosSoportados
 Configuracion.PortTCP = Server.PortTCP
 Configuracion.DireccionIPSMTP = Server.DireccionSMTP
 Configuracion.DireccionEMAILAdministrador = Server.DireccionEMAILAdministrador
 Configuracion.PermitirCrear = Server.PermitirCrearUsuarios

 ' **************************************************************
 ' Graba los Cambios de Configuracion
 ' **************************************************************
 GrabarConfiguracionArchivo
 
 ' **************************************************************
 ' Levanta el Sistema
 ' **************************************************************
 InicializarSistema
 MsgBox "The System was Succefcully ReStarted...", vbInformation, Configuracion.TituloVentanas
 EscribirEvento "The System was Succefcully ReStarted...", vbBlue
 
End Sub

Private Sub Lenguetas_Click(PreviousTab As Integer)
 
 CargarPantallasAutomaticas
  
End Sub

Private Sub RefrescoUsuarioActuales_Timer()

 If RefrescoUsuarioActuales Then
  RefrescoUsuarioActuales.Interval = Server.TiempoDeRefresco
  Varios.CargarListadoUsuariosActuales
 End If
 
End Sub
Private Sub ReiniciarSistema_Click()
Dim Respuesta As Integer

 ' **************************************************************
 ' Reinicia el Sistema
 ' **************************************************************
 ' Verifica que este seguro de Grabar y Restartear el Sistema
 Respuesta = MsgBox("¿Are you Sure to ReStart the System? (This Action Disconnect all User's and Change's Made)", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 If Respuesta = vbNo Then
  Exit Sub
 End If
  
 ' **************************************************************
 ' Detiene el Sistema
 ' **************************************************************
 DetenerSistema
 
 ' **************************************************************
 ' Deshace los Cambios
 ' **************************************************************
 Server.UbicacionBaseDeDatos = Configuracion.UbicacionBaseDeDatos
 Server.NombreDeLaBaseDeDatos = Configuracion.NombreDeLaBaseDeDatos
 Server.UsuariosSoportados = Configuracion.UsuariosSoportados
 Server.PortTCP = Configuracion.PortTCP
 Server.DireccionSMTP = Configuracion.DireccionIPSMTP
 Server.DireccionEMAILAdministrador = Configuracion.DireccionEMAILAdministrador
 Server.PermitirCrearUsuarios = Configuracion.PermitirCrear
  
 ' **************************************************************
 ' Levanta el Sistema
 ' **************************************************************
 InicializarSistema
 MsgBox "The System was Succefcully ReStarted...", vbInformation, Configuracion.TituloVentanas
 EscribirEvento "The System was Succefcully ReStarted...", vbBlue
 
End Sub

Private Sub SMTPMailSocket_DataArrival(ByVal bytesTotal As Long)
Dim Largo As Long

 ' **************************************************************
 ' Cuando llegan datos por el Winsock, los mismo
 ' Son guardados en Respuesta para su verificacion
 ' **************************************************************
 If SMTPMailSocket.State <> 6 Then
  Largo = bytesTotal
  SMTPMailSocket.GetData RespuestaWinsockSMTP
 End If
    
End Sub
Private Sub TCPSocket_Close(Index As Integer)
Dim Posicion As Integer

 ' **************************************************************
 ' Cierra el Socket...
 ' **************************************************************
 TCPSocket(Index).Close
  
 ' **************************************************************
 ' Busca el Usuario del Socket cerrado en la Matriz de Usuarios...
 ' **************************************************************
 Posicion = BuscarUserIDEnUsuarios(Sockets(Index).IDNumericoUsuario)
 
 ' **************************************************************
 ' Cambia los datos del usuario conectado en ese port...
 ' **************************************************************
 Usuarios(Posicion).EstadoActualNumero = 0
 Usuarios(Posicion).PortActual = 0
 Usuarios(Posicion).EstadoActualTexto = ""
 
 ' **************************************************************
 ' Cambia los datos de socket que se esta cerrando...
 ' **************************************************************
 Sockets(Index).EstadoDelPort = 0
 Sockets(Index).IDAliasUsuario = ""
 Sockets(Index).IDNumericoUsuario = 0
    
 ' **************************************************************
 ' Descuenta un Usuario de los Conectados
 ' **************************************************************
 Configuracion.UsuariosConectadosAlSistemas = Configuracion.UsuariosConectadosAlSistemas - 1
 Server.CantidadUsuariosLabel = Configuracion.UsuariosConectadosAlSistemas & " de " & Configuracion.UsuariosSoportados & "..."
 
 ' **************************************************************
 ' Evento de Sistema
 ' **************************************************************
 EscribirEvento "Desconection From " & TCPSocket(Index).RemoteHostIP, vbBlue
 
 ' **************************************************************
 ' Descarga el Socket...
 ' **************************************************************
 Unload TCPSocket(Index)
 
End Sub
Private Sub TCPSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim SocketDisponible As Integer

 ' **************************************************************
 ' Busca el Primer Socket Disponible para la Comunicacion
 ' Entrante...
 ' **************************************************************
 SocketDisponible = BuscarSocketDisponible
 ' Si el Valor de SocketDisponible es 0 es por que no hay socket
 ' disponibles...
 If SocketDisponible = 0 Then
  EscribirEvento "No More Available Connection's...", vbRed
  ' Sale sin dejar conectar
  Exit Sub
 End If
 
 ' **************************************************************
 ' Se redirecciona la Connection Request al Port Disponible
 ' **************************************************************
 If Index = 0 Then
  Load TCPSocket(SocketDisponible)
  ' Se envia la coneccion al socket Creado
  TCPSocket(SocketDisponible).Accept requestID
  ' *****
  ' Se Pasa el Socket a Estado Conectado... (Aunque en Realidad
  ' todavia no se valido al usuario...
  ' *****
  Sockets(SocketDisponible).EstadoDelPort = 1 ' Estado Logueando
  ' **************************************************************
  ' Agrega un Usuario a la cantidad de Conectados...
  ' **************************************************************
  Configuracion.UsuariosConectadosAlSistemas = Configuracion.UsuariosConectadosAlSistemas + 1
  Server.CantidadUsuariosLabel = Configuracion.UsuariosConectadosAlSistemas & " de " & Configuracion.UsuariosSoportados & "..."
 End If
  
 ' **************************************************************
 ' Evento de Sistema
 ' **************************************************************
 EscribirEvento "Connection Made From " & TCPSocket(Index).RemoteHostIP, vbBlue
 
End Sub
Private Sub TCPSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim DatosRecibidos, ComandoAccion, ComandoDatos, ResultadoComando As String
Dim PaqueteRecibido As Variant
Dim LargoDelPaquete As Integer
Dim ComandoCorrecto As Boolean

 ' **************************************************************
 ' Toma los datos que estan llegando al Socket
 ' **************************************************************
 TCPSocket(Index).GetData DatosRecibidos, vbString, bytesTotal
  
 ' **************************************************************
 ' Preprosesa el Paquete recibido a traves del Modulo (Recibir
 ' Paquete), en este punto se pueden hacer procesos de
 ' descompresion, etc.
 ' **************************************************************
 PaqueteRecibido = RecibirPaqueteTCP(DatosRecibidos)
 'Debug.Print PaqueteRecibido
 
 ' **************************************************************
 ' Todos los Paquetes como minimo deben medir 2 caracteres....
 ' Caracter 1:  Comando a Procesar
 ' Caracter 2:  Datos del Comando (Hasta el Final del Paquete)
 ' **************************************************************
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
 ' Si el Usuario No esta Logueado Solo Acepta Paquetes de
 ' Loguin (Comando 0)
 ' **************************************************************
 If Sockets(Index).IDNumericoUsuario = 0 And ComandoAccion <> "0" Then
  ' El usuario no esta logueado, por lo que se descarta la peticion
  ' y se escribe un evento
   EscribirEvento "Tray to Send From [" & TCPSocket(Index).RemoteHostIP & "] a Command not Logued-In... - Command [" & ComandoAccion & "]", vbRed
   Exit Sub
 End If
 
 ' **************************************************************
 ' Ejecuta el Comando Solicitado
 ' **************************************************************
 ComandoCorrecto = True
 Select Case ComandoAccion
  Case "0": ' Paquetes de Login
   ResultadoComando = ComandoAccion_0(ComandoDatos, Index)
  Case "1": ' Solicitar Lista de Amigos
   ResultadoComando = ComandoAccion_1(ComandoDatos, Index)
  Case "2": ' Cambio en Lista de Amigos
   ResultadoComando = ComandoAccion_2(ComandoDatos, Index)
  Case "3": ' Intercambio de Paquetes (El Server no Tiene Intervencion)
   ResultadoComando = ComandoAccion_3(ComandoDatos, Index)
  Case "4": ' Intercambiar Mensaje
   ResultadoComando = ComandoAccion_4(ComandoDatos, Index)
  Case "5": ' Mensaje Offline
   ResultadoComando = ComandoAccion_5(ComandoDatos, Index)
  Case Else:
   ComandoCorrecto = False
 End Select
 
 ' **************************************************************
 ' Si la Funcion solicitada no es correcta pone un aviso via
 ' via log...
 ' **************************************************************
 If ComandoCorrecto = False Then
  EscribirEvento "Tray to Send From [" & TCPSocket(Index).RemoteHostIP & "] a Command not Logued-In... - Command [" & ComandoAccion & "]", vbRed
 End If
 
 ' **************************************************************
 ' Si la respuesta del Comando Ejecutado es Nula quiere decir
 ' que se ejecuto Correctamente... Sino se muestra en el LOG
 ' el error...
 ' **************************************************************
 If ResultadoComando <> "" Then
  EscribirEvento ResultadoComando, vbRed
 End If
  
End Sub

Private Sub TCPSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim Posicion As Integer

 ' **************************************************************
 ' Cuando hay un error, se desconecta el Socket...
 ' **************************************************************
 TCPSocket(Index).Close
 Unload TCPSocket(Index)
 ' Se cambia el Estado de Dicho Socket
 Sockets(Index).EstadoDelPort = 0 ' Estado Desconectado
 
 ' **************************************************************
 ' Descuenta un Usuario de los Conectados
 ' **************************************************************
 Configuracion.UsuariosConectadosAlSistemas = Configuracion.UsuariosConectadosAlSistemas - 1
 Server.CantidadUsuariosLabel = Configuracion.UsuariosConectadosAlSistemas & " de " & Configuracion.UsuariosSoportados & "..."
 
 ' **************************************************************
 ' Se Cambia los datos del usuario conectado en este port...
 ' **************************************************************
 If Sockets(Index).IDNumericoUsuario <> 0 Then
  ' Cambiar el Estado del Usuario en la Matriz de Usuarios
  Posicion = Varios.BuscarUserIDEnUsuarios(Sockets(Index).IDNumericoUsuario)
  With Usuarios(Posicion)
   .EstadoActualNumero = 0
   .EstadoActualTexto = ""
   .PortActual = 0
  End With
 End If
 
 ' **************************************************************
 ' Si el Socket es el 0 Entonces lo vuelve a cargar para que
 ' el sistema siga funcionando...
 ' **************************************************************
 If Index = 0 Then
   Load TCPSocket(0)
   TCPSocket(0).Listen
 End If
 
 ' **************************************************************
 ' Evento de Sistema
 ' **************************************************************
 EscribirEvento "Error Nro. " & Number & ", " & Description & ". Socket [" & Index & "]", vbRed
 
 
End Sub

Private Sub UserManBorrarUsuario_Click()
Dim Posicion As Integer
Dim AliasUsuario As String

 If Trim(UsuariosRegistrados.Text) = "" Then Exit Sub ' Nada para hacer
  
 ' **************************************************************
 ' Separa el IDDeUsuario
 ' **************************************************************
 Posicion = InStr(1, UsuariosRegistrados.Text, "- [")
 If Posicion > 0 Then ' Separa el IDAlias del Usuario...
  AliasUsuario = Mid$(UsuariosRegistrados.Text, 1, Posicion - 1)
 End If

 ' **************************************************************
 ' Esta Seguro?
 ' **************************************************************
 Posicion = MsgBox("¿Are You Sure to Delete the User [" & AliasUsuario & "]?...", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
 If Posicion = vbNo Then
  Exit Sub
 End If

 ' **************************************************************
 ' Buscar Usuario...
 ' **************************************************************
 Posicion = BuscarUsuarioAliasEnUsuarios(AliasUsuario)
 If Posicion = 0 Then Exit Sub ' Nada por Hacer...
 
 ' **************************************************************
 ' Verifica si esta conectado...
 ' **************************************************************
 If Usuarios(Posicion).PortActual <> 0 Then
  TCPSocket(CInt(Posicion)).Close
  TCPSocket_Close (CInt(Posicion))
 End If
 
 ' **************************************************************
 ' Borrar el Usuario (De la Base)
 ' **************************************************************
 With Usuarios(Posicion)
  .ApellidoYNombre = ""
  .DireccionDeEmail = ""
  .Edad = ""
  .EstadoCivil = ""
  .FechaDeNacimiento = ""
  .Humor = ""
  .IDAliasUsuario = ""
  .Intencion = ""
  .ListadoDeAmigos = ""
  .MensajesOffline = ""
  .Ocupacion = ""
  .OtraInfo = ""
  .Password = ""
  .PortActual = 0
  .Sexo = ""
  .SigNo = ""
  .Telefono = ""
  .UbicacionGeografica = ""
  .UsuarioBloqueado = False
  .UltimoLogueo = ""
 End With
 BaseDeDatos.GrabarModificacionesUsuario Posicion, True
  
  
 ' **************************************************************
 ' Recarga el Listado de Usuarios...
 ' **************************************************************
 CargarListadoDeUsuarioRegistrado
  
End Sub
Private Sub UserManCrearUsuario_Click()
Dim Contador, Disponible As Integer

 If Me.UserManCrearUsuario.Caption = "Add User..." Then
  ' **************************************************************
  ' Prepara el Formulario...
  ' **************************************************************
  Me.BlanquearCamposUserManager
  Me.UserManCrearUsuario.Caption = "Add..."
  Me.UserManDescartarCambios.Caption = "Cancel..."
  Me.UserManBorrarUsuario.Enabled = False
  Me.UserManGrabarCambios.Enabled = False
  Me.UsuariosRegistrados.Enabled = False
  Me.User_IDAliasUsuario.Enabled = True
  IndiceActual = Me.UsuariosRegistrados.ListIndex
  Exit Sub
 End If
  
 ' **************************************************************
 ' Busca el Primer Disponible... (Si no hay Disponible Avisa)...
 ' **************************************************************
 Disponible = 0
 For Contador = 1 To Configuracion.UsuariosSoportados
  If Trim(Usuarios(Contador).IDAliasUsuario) = "" Then
   Disponible = Contador
   Exit For
  End If
 Next
 If Disponible = 0 Then ' No hay Disponibles...
  MsgBox "Sorry, The User [" & Me.User_IDAliasUsuario & "] can't be Added Beacause no more Availables User's...", vbCritical, Configuracion.TituloVentanas
  FormularioNormal
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verificar que no sea un Usuario Duplicado...
 ' **************************************************************
 Me.UserManDescartarCambios.Enabled = False
 Me.UserManCrearUsuario.Enabled = False
 For Contador = 1 To Configuracion.UsuariosSoportados
  If UCase(Trim(Usuarios(Contador).IDAliasUsuario)) = UCase(Trim(Me.User_IDAliasUsuario)) Then
   MsgBox "The User Alias [" & Trim(Me.User_IDAliasUsuario) & "] already Exist, Please Enter Another...", vbCritical + vbOKOnly, Configuracion.TituloVentanas
   Me.UserManDescartarCambios.Enabled = True
   Me.UserManCrearUsuario.Enabled = True
   Exit Sub
  End If
 Next
 
 ' **************************************************************
 ' Verificar que los campos necesarios existan...
 ' **************************************************************
 Dim MensajeError As String
 MensajeError = ""
 If Trim(Me.User_IDAliasUsuario) = "" Then MensajeError = "You Must Enter a Valid User Alias..."
 If Trim(Me.User_ApellidoYNombre) = "" Then MensajeError = "You Must Enter a Valid Name..."
 If Trim(Me.User_DireccionDeEmail) = "" Then MensajeError = "You Must Enter a Valid E-Mail Direction..."
 If Trim(Me.User_Password) = "" Then
  Me.User_Password = Varios.GenerarPassword(8)
 End If
 If MensajeError <> "" Then
  MsgBox MensajeError, vbCritical, Configuracion.TituloVentanas
  Me.UserManDescartarCambios.Enabled = True
  Me.UserManCrearUsuario.Enabled = True
  Exit Sub
 End If
 
 ' **************************************************************
 ' Comienza a Grabar...
 ' **************************************************************
 Me.GrabarCambioUsuario Disponible ', Varios.Mensaje_Bienvenida
 
 
 ' **************************************************************
 ' Envia un EMail con Usuario y Password...
 ' **************************************************************
 Dim Respuesta As Boolean
 Respuesta = Varios.EnviarPasswordAUsuario(Trim(Me.User_IDAliasUsuario), Trim(Me.User_Password), Trim(Me.User_DireccionDeEmail), Trim(Me.User_ApellidoYNombre))
 If Respuesta = False Then
  MsgBox "An Error Ocurred When Try to Send the Password to [" & Trim(Me.User_IDAliasUsuario) & "] to [" & Trim(Me.User_DireccionDeEmail) & "]...", vbOKOnly + vbCritical, Configuracion.TituloVentanas
 End If
 
 ' **************************************************************
 ' Confirma el Ok...
 ' **************************************************************
 MsgBox "The User [" & Me.User_IDAliasUsuario & "] was Added...", vbInformation, Configuracion.TituloVentanas
  
 ' **************************************************************
 ' Desbloquea los botones...
 ' **************************************************************
 Me.UserManDescartarCambios.Enabled = True
 Me.UserManCrearUsuario.Enabled = True
  
 ' **************************************************************
 ' Blanquea los Campos...
 ' **************************************************************
 Varios.CargarListadoDeUsuarioRegistrado
 Me.BlanquearCamposUserManager
 
End Sub
Public Sub FormularioNormal()
Dim Contador As Integer
Dim AliasUsuario As String

 Me.UserManCrearUsuario.Caption = "Add User..."
 Me.UserManDescartarCambios.Caption = "Discard Change's..."
 Me.UserManBorrarUsuario.Enabled = True
 Me.UserManGrabarCambios.Enabled = True
 Me.UsuariosRegistrados.Enabled = True
 Me.User_IDAliasUsuario.Enabled = False
 If Me.UsuariosRegistrados.ListCount = 0 Then
  Me.BlanquearCamposUserManager
  Exit Sub
 End If
 If IndiceActual <> Null Or Me.UsuariosRegistrados.ListCount <> 0 Then
  ' **************************************************************
  ' Se posiciona en el Ultimo Usuario...
  ' **************************************************************
   
   Me.UsuariosRegistrados.ListIndex = IndiceActual
  ' Separa el IDDeUsuario
  Contador = InStr(1, Server.UsuariosRegistrados.Text, "- [")
  If Contador > 0 Then ' Separa el IDAlias del Usuario...
   AliasUsuario = Mid$(Server.UsuariosRegistrados.Text, 1, Contador - 1)
  End If
  Server.CargarUsuario_UserManager (AliasUsuario)
 End If

End Sub
Private Sub UserManDescartarCambios_Click()
Dim Posicion As Long

 ' **************************************************************
 ' Descartar Cambios...
 ' **************************************************************
 If Me.UserManCrearUsuario.Caption <> "Add..." Then
  ' **************************************************************
  ' Hubo Cambios...
  ' **************************************************************
  If Me.VerificarCambioDatos = False Then
   Exit Sub
  End If
  
  ' **************************************************************
  ' Esta Seguro?
  ' **************************************************************
  Posicion = MsgBox("¿Are You Sure to Discard The Change's?...", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
  If Posicion = vbYes Then
   Me.CargarUsuario_UserManager (Trim(Me.User_IDAliasUsuario))
  End If
  Exit Sub
 End If
 
 If Me.UserManCrearUsuario.Caption = "Add..." Then
  FormularioNormal
  Exit Sub
 End If
 
 
End Sub

Private Sub UserManGrabarCambios_Click()
 
 ' **************************************************************
 ' Verificar que los campos necesarios existan...
 ' **************************************************************
 Dim MensajeError As String
 MensajeError = ""
 If Trim(Me.User_IDAliasUsuario) = "" Then MensajeError = "You Must Enter a Valid User Alias..."
 If Trim(Me.User_ApellidoYNombre) = "" Then MensajeError = "You Must Enter a Valid Name..."
 If Trim(Me.User_DireccionDeEmail) = "" Then MensajeError = "You Must Enter a Valid E-Mail Direction..."
 If Trim(Me.User_Password) = "" Then MensajeError = "You Must Enter a Valid Password..."
 If MensajeError <> "" Then
  MsgBox MensajeError, vbCritical, Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Verificar Cambio de Usuario...
 ' **************************************************************
 If Me.VerificarCambioDatos = False Then
  MsgBox "You Are Not Change the User Profile...", vbInformation + vbOKOnly, Configuracion.TituloVentanas
  Exit Sub
 End If
 
 ' **************************************************************
 ' Graba los Cambios...
 ' **************************************************************
 Me.GrabarCambioUsuario ' Graba los Cambios...
 
End Sub

Private Sub UsuariosRegistrados_Click()
Dim AliasUsuario As String
Dim Posicion As Long

 If Trim(UsuariosRegistrados.Text) = "" Then Exit Sub ' Nada para hacer
  
 ' **************************************************************
 ' Separa el IDDeUsuario
 ' **************************************************************
 Posicion = InStr(1, UsuariosRegistrados.Text, "- [")
 If Posicion > 0 Then ' Separa el IDAlias del Usuario...
  AliasUsuario = Mid$(UsuariosRegistrados.Text, 1, Posicion - 1)
 End If
 
 ' **************************************************************
 ' Verificar Cambio de Usuario...
 ' **************************************************************
 If Me.VerificarCambioDatos Then
  ' Se hicieron Cambios...
  Posicion = MsgBox("You Make Change's to User [" & Trim(Me.User_IDAliasUsuario) & "]... ¿Do you Like to Save the Change's?", vbQuestion + vbYesNo, Configuracion.TituloVentanas)
  If Posicion = vbYes Then
   Me.GrabarCambioUsuario ' Graba los Cambios...
  End If
 End If
 
 ' **************************************************************
 ' Carga los Datos del Nuevo Usuario...
 ' **************************************************************
 CargarUsuario_UserManager (AliasUsuario)
 
End Sub
Public Sub CargarUsuario_UserManager(AliasUsuario As String)
Dim Posicion As Integer

 If Trim(AliasUsuario) = "" Then Exit Sub ' Nada por Hacer...
 
 ' **************************************************************
 ' Buscar el Numero de Usuario...
 ' **************************************************************
 Posicion = BuscarUsuarioAliasEnUsuarios(AliasUsuario)
 If Posicion = 0 Then Exit Sub ' Nada por Hacer...
 
 ' **************************************************************
 ' Empezar a Procesar el Usuario...
 ' **************************************************************
 With Usuarios(Posicion)
  Me.User_IDAliasUsuario = Trim(.IDAliasUsuario)
  Me.User_ApellidoYNombre = Trim(.ApellidoYNombre)
  Me.User_DireccionDeEmail = Trim(.DireccionDeEmail)
  Me.User_FechadeNacimiento = Trim(.FechaDeNacimiento)
  Me.User_Edad = Trim(.Edad)
  'Me.User_EstadoCivil = Trim(.EstadoCivil)
  If .EstadoCivil = "C" Then Me.User_EstadoCivil = "Married"
  If .EstadoCivil = "S" Then Me.User_EstadoCivil = "Single"
  If .EstadoCivil = "V" Then Me.User_EstadoCivil = "Widowed"
  If .EstadoCivil = "D" Then Me.User_EstadoCivil = "Divorced"
  Me.User_Humor = Trim(.Humor)
  Me.User_Intencion = Trim(.Intencion)
  Me.User_Ocupacion = Trim(.Ocupacion)
  Me.User_OtraInfo = Trim(.OtraInfo)
  If UCase(Trim(.Sexo)) = "M" Then Me.User_Sexo = "Male"
  If UCase(Trim(.Sexo)) = "F" Then Me.User_Sexo = "Female"
  Me.User_Telefono = Trim(.Telefono)
  Select Case Trim(.SigNo)
   Case "Capricornio"
    Me.User_Signo = "Capricornia"
   Case "Acuario"
    Me.User_Signo = "Aquaria"
   Case "Picis"
    Me.User_Signo = "Pisis"
   Case "Aries"
    Me.User_Signo = "Aries"
   Case "Tauro"
    Me.User_Signo = "Taurus"
   Case "Geminis"
    Me.User_Signo = "Gemini"
   Case "Cancer"
    Me.User_Signo = "Cancer"
   Case "Leo"
    Me.User_Signo = "Leo"
   Case "Virgo"
    Me.User_Signo = "Virgo"
   Case "Libra"
    Me.User_Signo = "Libra"
   Case "Escorpio"
    Me.User_Signo = "Scorpio"
   Case "Sagitario"
    Me.User_Signo = "Sagitarius"
  End Select
  ' Me.User_Signo = Trim(.Signo)
  Me.User_UbicacionGeografica = Trim(.UbicacionGeografica)
  Me.User_UsuarioBloqueado = .UsuarioBloqueado
  Me.User_Password = Trim(.Password)
 End With
 
End Sub
Public Function VerificarCambioDatos() As Boolean
Dim Posicion As Integer
Dim Cambio As Boolean

 ' **************************************************************
 ' Nada por Hacer...
 ' **************************************************************
 If Trim(Me.User_IDAliasUsuario) = "" Then
  VerificarCambioDatos = False
  Exit Function
 End If
 
 ' **************************************************************
 ' Buscar el Numero de Usuario...
 ' **************************************************************
 Posicion = BuscarUsuarioAliasEnUsuarios(Me.User_IDAliasUsuario)
 If Posicion = 0 Then
  VerificarCambioDatos = False
  Exit Function ' Nada por Hacer...
 End If
 
 ' **************************************************************
 ' Verificar si se Cambio Algo...
 ' **************************************************************
 Cambio = False
 With Usuarios(Posicion)
  If Trim(Me.User_IDAliasUsuario) <> Trim(.IDAliasUsuario) Then Cambio = True
  If Trim(Me.User_ApellidoYNombre) <> Trim(.ApellidoYNombre) Then Cambio = True
  If Trim(Me.User_DireccionDeEmail) <> Trim(.DireccionDeEmail) Then Cambio = True
  If Trim(Me.User_FechadeNacimiento) <> Trim(.FechaDeNacimiento) Then Cambio = True
  If Trim(Me.User_Edad) <> Trim(.Edad) Then Cambio = True
  ' Arregla Estado Civil...
  Dim EstadoTemp As String
  If Me.User_EstadoCivil = "Married" Then EstadoTemp = "C"
  If Me.User_EstadoCivil = "Single" Then EstadoTemp = "S"
  If Me.User_EstadoCivil = "Widowed" Then EstadoTemp = "V"
  If Me.User_EstadoCivil = "Divorced" Then EstadoTemp = "D"
  If UCase(Trim(EstadoTemp)) <> UCase(Trim(.EstadoCivil)) Then Cambio = True
  ' ***
  If Trim(Me.User_Humor) <> Trim(.Humor) Then Cambio = True
  If Trim(Me.User_Intencion) <> Trim(.Intencion) Then Cambio = True
  If Trim(Me.User_Ocupacion) <> Trim(.Ocupacion) Then Cambio = True
  If Trim(Me.User_OtraInfo) <> Trim(.OtraInfo) Then Cambio = True
  ' Arregla el Sexo...
  If UCase(Mid$(Trim(Me.User_Sexo), 1, 1)) <> UCase(Trim(.Sexo)) Then Cambio = True
  If Trim(Me.User_Telefono) <> Trim(.Telefono) Then Cambio = True
  
  ' Arregla el Signo...
  Dim SigNo As String
  Select Case Trim(Me.User_Signo)
   Case "Capricornia"
    SigNo = "Capricornio"
   Case "Aquaria"
    SigNo = "Acuario"
   Case "Pisis"
    SigNo = "Picis"
   Case "Aries"
    SigNo = "Aries"
   Case "Taurus"
    SigNo = "Tauro"
   Case "Gemini"
    SigNo = "Geminis"
   Case "Cancer"
    SigNo = "Cancer"
   Case "Leo"
    SigNo = "Leo"
   Case "Virgo"
    SigNo = "Virgo"
   Case "Libra"
    SigNo = "Libra"
   Case "Scorpio"
    SigNo = "Escorpio"
   Case "Sagitarius"
    SigNo = "Sagitario"
  End Select
  If Trim(SigNo) <> Trim(.SigNo) Then Cambio = True
  If Trim(Me.User_UbicacionGeografica) <> Trim(.UbicacionGeografica) Then Cambio = True
  If Me.User_UsuarioBloqueado <> .UsuarioBloqueado Then Cambio = True
  If Trim(Me.User_Password) <> Trim(.Password) Then Cambio = True
 End With
 VerificarCambioDatos = Cambio
 
End Function
Public Function GrabarCambioUsuario(Optional NumeroPosicion As Integer, Optional MensajeOffLine As String) As Boolean
Dim Posicion As Integer
Dim Cambio As Boolean
Dim Nuevo As Boolean

 ' **************************************************************
 ' Se usa como bandera...
 ' **************************************************************
 Posicion = 0
 
 ' **************************************************************
 ' Si se pasa un Numero lo usa para identificar donde debe Grabr
 ' **************************************************************
 If IsNumeric(NumeroPosicion) Then
  If NumeroPosicion <> 0 Then
   Posicion = NumeroPosicion
   Nuevo = True
  End If
 End If
 
 If Posicion = 0 Then
  ' **************************************************************
  ' Nada por Hacer...
  ' **************************************************************
  If Trim(Me.User_IDAliasUsuario) = "" Then
   GrabarCambioUsuario = False
   Exit Function
  End If
  ' **************************************************************
  ' Buscar el Numero de Usuario...
  ' **************************************************************
  Posicion = BuscarUsuarioAliasEnUsuarios(Me.User_IDAliasUsuario)
  If Posicion = 0 Then
   GrabarCambioUsuario = False
   Exit Function ' Nada por Hacer...
  End If
 End If
 
 ' **************************************************************
 ' Graba los Cambios...
 ' **************************************************************
 With Usuarios(Posicion)
  If Nuevo Then .IDAliasUsuario = Trim(Me.User_IDAliasUsuario)
  .ApellidoYNombre = Me.User_ApellidoYNombre
  .DireccionDeEmail = Me.User_DireccionDeEmail
  .FechaDeNacimiento = Me.User_FechadeNacimiento
  .Edad = Me.User_Edad
  If Me.User_EstadoCivil = "Married" Then .EstadoCivil = "C"
  If Me.User_EstadoCivil = "Single" Then .EstadoCivil = "S"
  If Me.User_EstadoCivil = "Widowed" Then .EstadoCivil = "V"
  If Me.User_EstadoCivil = "Divorced" Then .EstadoCivil = "D"
  .Humor = Me.User_Humor
  .Intencion = Me.User_Intencion
  .Ocupacion = Me.User_Ocupacion
  .OtraInfo = Me.User_OtraInfo
  If UCase(Trim(Me.User_Sexo)) = UCase(Trim("Male")) Then .Sexo = "M"
  If UCase(Trim(Me.User_Sexo)) = UCase(Trim("Female")) Then .Sexo = "F"
  .Telefono = Me.User_Telefono
  '.Signo = Me.User_Signo
  Select Case Trim(Me.User_Signo)
   Case "Capricornia"
    .SigNo = "Capricornio"
   Case "Aquaria"
    .SigNo = "Acuario"
   Case "Pisis"
    .SigNo = "Picis"
   Case "Aries"
    .SigNo = "Aries"
   Case "Taurus"
    .SigNo = "Tauro"
   Case "Gemini"
    .SigNo = "Geminis"
   Case "Cancer"
    .SigNo = "Cancer"
   Case "Leo"
    .SigNo = "Leo"
   Case "Virgo"
    .SigNo = "Virgo"
   Case "Libra"
    .SigNo = "Libra"
   Case "Scorpio"
    .SigNo = "Escorpio"
   Case "Sagitarius"
    .SigNo = "Sagitario"
  End Select
  .UbicacionGeografica = Me.User_UbicacionGeografica
  .UsuarioBloqueado = Me.User_UsuarioBloqueado
  .Password = Me.User_Password
  If Trim(MensajeOffLine) <> "" Then
   .MensajesOffline = MensajeOffLine
  End If
 End With
 BaseDeDatos.GrabarModificacionesUsuario (Posicion)
 GrabarCambioUsuario = True
 
End Function
Public Sub BlanquearCamposUserManager()

  Me.User_IDAliasUsuario = ""
  Me.User_ApellidoYNombre = ""
  Me.User_DireccionDeEmail = ""
  Me.User_FechadeNacimiento = ""
  Me.User_Edad = ""
  Me.User_EstadoCivil = "Married"
  Me.User_Humor = ""
  Me.User_Intencion = ""
  Me.User_Ocupacion = ""
  Me.User_OtraInfo = ""
  Me.User_Sexo = "Male"
  Me.User_Telefono = ""
  Me.User_Signo = "Capricornia"
  Me.User_UbicacionGeografica = ""
  Me.User_UsuarioBloqueado = False
  Me.User_Password = ""
  
End Sub
