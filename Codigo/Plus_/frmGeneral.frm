VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmGeneral 
   BackColor       =   &H00000000&
   Caption         =   "Argentum Online - GSS - ""eL33T"""
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   -600
   ClientWidth     =   11070
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Tag             =   "v0.12b3 fix-1 - ""T-Fire"""
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   960
      Top             =   2160
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   480
      Top             =   2160
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1535
      ButtonWidth     =   2117
      ButtonHeight    =   1376
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "Iconos"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Control"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&NPC Debug"
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Socket's"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Trafico"
            Object.Tag             =   ""
            ImageIndex      =   41
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Configuración"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Alertas"
            Object.Tag             =   ""
            ImageIndex      =   14
            Object.Width           =   6e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CommandButton sendMSGx 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "&Enviar mensaje"
         Default         =   -1  'True
         Height          =   400
         Left            =   9570
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   345
         Width           =   1335
      End
      Begin VB.ListBox Dirigido 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C000&
         Height          =   420
         ItemData        =   "frmGeneral.frx":1042
         Left            =   8040
         List            =   "frmGeneral.frx":104C
         TabIndex        =   5
         Top             =   340
         Width           =   1455
      End
      Begin VB.ListBox Por 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C000&
         Height          =   420
         ItemData        =   "frmGeneral.frx":106F
         Left            =   7200
         List            =   "frmGeneral.frx":1079
         TabIndex        =   4
         Top             =   340
         Width           =   735
      End
      Begin VB.TextBox Mensaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   7200
         MaxLength       =   1024
         TabIndex        =   3
         Top             =   40
         Width           =   3735
      End
   End
   Begin ComctlLib.ProgressBar ProG1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   7470
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer NoMain 
      Interval        =   1
      Left            =   3840
      Top             =   2760
   End
   Begin ComctlLib.StatusBar Publicidad 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   7725
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   "GS Server AO v0.14a - ""eL33T"" - Programado por ^[GS]^ - E-mail: gshaxor@gmail.com - Versión Actual: 14 de Agosto del 2005"
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer NpcAtaca 
      Interval        =   4000
      Left            =   3840
      Top             =   2160
   End
   Begin VB.Timer FX 
      Interval        =   200
      Left            =   2400
      Top             =   2160
   End
   Begin VB.Timer tLluviaEvent 
      Interval        =   60000
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   1920
      Top             =   2160
   End
   Begin VB.Timer CmdExec 
      Interval        =   1
      Left            =   1440
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1920
      Top             =   1440
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2880
      Top             =   1440
   End
   Begin VB.Timer tLluvia 
      Interval        =   500
      Left            =   1440
      Top             =   1440
   End
   Begin VB.Timer tPiqueteC 
      Interval        =   6000
      Left            =   3360
      Top             =   1440
   End
   Begin VB.Timer tDeRepetir 
      Interval        =   1200
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer TIMER_AI 
      Interval        =   100
      Left            =   3360
      Top             =   2160
   End
   Begin VB.Timer KillLog 
      Interval        =   60000
      Left            =   2880
      Top             =   2160
   End
   Begin VB.Timer Auditoria 
      Interval        =   1000
      Left            =   960
      Top             =   1440
   End
   Begin VB.Timer TimerCartelito 
      Interval        =   5000
      Left            =   480
      Top             =   1440
   End
   Begin ComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Programado por ^[GS]^"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Slave 
      Index           =   0
      Left            =   2040
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   111
   End
   Begin MSWinsockLib.Winsock master 
      Left            =   2520
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   111
   End
   Begin ComctlLib.ImageList Iconos 
      Left            =   4200
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16384
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   48
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":108F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":13A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":16C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":19DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":1CF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":2011
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":232B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":2645
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":295F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":2C79
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":2F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":32AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":35C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":38E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":3BFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":3F15
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":422F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":4549
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":4863
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":4B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":4E97
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":51B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":54CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":57E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":5AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":5E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":6133
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":644D
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":6767
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":6A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":6D9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":70B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":73CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":76E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":7A03
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":7D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":8037
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":8351
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":866B
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":8985
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":8C9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":8FB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":92D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":95ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":9907
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":9C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":9F3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGeneral.frx":A255
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuOcultar 
         Caption         =   "&Ocultar en el Systray"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrarCorrectamente 
         Caption         =   "&Cerrar el Servidor correctamente (Backup + Personajes + Clanes)"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "&Cerrar el Servidor"
      End
      Begin VB.Menu end1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ventanas"
      Begin VB.Menu mnuPaneldeControl 
         Caption         =   "&Panel de Control"
      End
      Begin VB.Menu mnuEscaneadordePjs 
         Caption         =   "&Escaneador de PJs"
      End
      Begin VB.Menu mnuConfiguración 
         Caption         =   "&Configuración"
      End
      Begin VB.Menu mnuValidClien 
         Caption         =   "&Validadación de Cliente Propio"
      End
      Begin VB.Menu linX12 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuDepuracion 
      Caption         =   "&Depuración"
      Begin VB.Menu mnuSocketDebug 
         Caption         =   "&Socket's"
      End
      Begin VB.Menu mnuTrafico 
         Caption         =   "&Trafico"
      End
      Begin VB.Menu mnuNPCDebug 
         Caption         =   "&NPC Debug"
      End
      Begin VB.Menu mnuAlertas 
         Caption         =   "&Alertas e Informaciones"
      End
      Begin VB.Menu end2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuActualizar 
      Caption         =   "&Actualizar"
      Begin VB.Menu mnuGuardiasPos 
         Caption         =   "... posición de los &Guardias"
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "... &NPC's"
      End
      Begin VB.Menu mnuNPCsR 
         Caption         =   "... N&PC's (Recargandolos)"
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "... &Objetos"
      End
      Begin VB.Menu mnuHechizos 
         Caption         =   "... &Hechizos"
      End
      Begin VB.Menu end3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "&Acciones"
      Begin VB.Menu mnuMOTD 
         Caption         =   "Recargar Mensaje del &Dia (MOTD)"
      End
      Begin VB.Menu mnuServidorINI 
         Caption         =   "Recargar &Server.ini (Reinicia los Sockets)"
      End
      Begin VB.Menu mnuReloadOpciones 
         Caption         =   "Recargar &Opciones.ini y Estadisticas.ini"
      End
      Begin VB.Menu mnuNombresProhibidos 
         Caption         =   "Recargar Nombres &Prohibidos"
      End
      Begin VB.Menu mnuReLoadSpawn 
         Caption         =   "Recargar lista de Spa&wn (Invokar.dat)"
      End
      Begin VB.Menu mnuUnBanIP 
         Caption         =   "&UNBAN todos los IP's"
      End
      Begin VB.Menu mnuUnBAN 
         Caption         =   "&UNBAN a todos"
      End
      Begin VB.Menu end4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarPersonajesYClanes 
         Caption         =   "&Guardar Personajes y Clanes"
      End
      Begin VB.Menu mnuHacerBackUp 
         Caption         =   "&Hacer un BackUp"
      End
      Begin VB.Menu mnuCargarDesdeBackUp 
         Caption         =   "&Cargar desde BackUp"
      End
   End
   Begin VB.Menu espacio1 
      Caption         =   "- - - - -"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuCreditos 
      Caption         =   "&Creditos"
      Begin VB.Menu mnuBuscaActualizar 
         Caption         =   "&Buscar Actualizaciones..."
      End
      Begin VB.Menu lin4 
         Caption         =   "-"
      End
      Begin VB.Menu cred1 
         Caption         =   "&Programado por ^[GS]^"
      End
      Begin VB.Menu cred3 
         Caption         =   "&E-mail: gshaxor@gmail.com"
      End
      Begin VB.Menu mnuWEB 
         Caption         =   "&Web: www.gs-zone.com.ar"
      End
      Begin VB.Menu lin5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu espacio2 
      Caption         =   "- - - - -"
   End
   Begin VB.Menu mnuMan 
      Caption         =   "Mantenimiento en..."
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar..."
      End
      Begin VB.Menu Linx11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActualizarPOP 
         Caption         =   "&Actualizar"
         Begin VB.Menu mnuPOPP 
            Caption         =   "... posición de los &Guardias"
         End
         Begin VB.Menu mnuPOPN 
            Caption         =   "... &NPC's"
         End
         Begin VB.Menu mnuPOPNR 
            Caption         =   "... N&PC's (Recargandolos)"
         End
         Begin VB.Menu menuPOPO 
            Caption         =   "... &Objetos"
         End
         Begin VB.Menu mnuPOPH 
            Caption         =   "... &Hechizos"
         End
      End
      Begin VB.Menu linX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPOPUNBANIP 
         Caption         =   "&UNBAN todos los IP"
      End
      Begin VB.Menu mnuGuardarPOP 
         Caption         =   "&Guardar Personajes y Clanes"
      End
      Begin VB.Menu mnuHacerBackPOP 
         Caption         =   "&Hacer un BackUp"
      End
      Begin VB.Menu linX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrarPop 
         Caption         =   "&Cerrar el Servidor"
      End
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Icono As Object
'Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer


Private Sub cred1_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText "gshaxor@gmail.com"
Call FrmMensajes.msg("Nota", "E-mail copiado.")
End Sub

Private Sub cred3_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText "gshaxor@gmail.com"
Call FrmMensajes.msg("Nota", "E-mail copiado.")
End Sub

Private Sub MDIForm_Activate()
' XP Stylus

End Sub

Private Sub MDIForm_Load()
On Error GoTo fallo
' **** NO CLOSE ****
'Dim MenuSistema%, Res%
'MenuSistema% = GetSystemMenu(hWnd, 0)
'Res% = RemoveMenu(MenuSistema%, 6, MF_BYPOSITION)
' **** NO CLOSE ****
Por.ListIndex = 0
Dirigido.ListIndex = 0
frmG_Main.Hide
frmG_Alertas.Hide
Exit Sub
fallo:
'MsgBox "Error cargando frmGeneral"
End Sub

Function PonerSysTray()
' [GS] Viejo systray
On Error GoTo fallo
frmGeneral.Hide
frmGeneral.Visible = False
Set Icono = frmGeneral.Icon
AddIcon frmCargando, Icono.Handle, Icono, "Argentum Online Server"
Exit Function
fallo:
If Err.Number = 28 Then Exit Function   ' Error de Stack, todo BIEN :D
Call FrmMensajes.msg("Error", "Error Poniendo en SysTray")
' [/GS]

End Function

Function QuitarSysTray()
' [GS] Viejo systray :P
On Error GoTo fallo
delIcon Icono.Handle
frmGeneral.Show
frmGeneral.Visible = True
Exit Function
fallo:
Call FrmMensajes.msg("Error", "Error quitando del SysTray")
' [/GS]

End Function


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'MsgBox Button

'If mx = 517 Or mx = 518 Then
'    PopupMenu Me.mnuPopupMenu
'End If
' Dim Message As Long
'    Message = x / Screen.TwipsPerPixelX
'    Select Case Message
'    Case WM_RBUTTONUPx
'        PopupMenu Me.mnuPopupMenu
'    Case Else
'        MsgBox Message
'    End Select


End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If haciendoBK = True Then
    Call FrmMensajes.msg("Alerta", "No puedes cerrar mientras realizas un backup.")
    Cancel = True
    Exit Sub
End If
' Guardar el pozo de la loteria
Call WriteVar(IniPath & "Estadisticas.ini", "LOTERIA", "Pozo_Loteria", val(Pozo_Loteria))
DoEvents

' [GS] Nuevo :D
Call WriteVar(IniPath & "Server.ini", "SEGURIDAD", "Funcionando", 0)
' [/GS]

If mnuCerrarCorrectamente.Checked = True Then Exit Sub
If mnuCerrar.Checked = True Then Exit Sub

Call GuardarUsuarios
Call SaveGuildsDB

DoEvents
If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    mnuCerrar.Checked = True
    DoEvents
    Dim f
    For Each f In Forms
        Unload f
    Next
Else
    Cancel = True
End If

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then
    PonerSysTray
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next



QuitarSysTray

#If UsarAPI Then
Call LimpiaWsApi
#Else
Socket1.Cleanup
#End If

Call DescargaNpcsDat

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

' Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " Servidor Cerrado. (Tiempo: " & Horas & " hs - " & MinsRunning & " Minutos.)"
Close #N

End

End Sub

Private Sub menuPOPO_Click()
Call mnuObjetos_Click
End Sub

Private Sub mnuAlertas_Click()
On Error Resume Next
        frmG_Alertas.Show
        frmG_Alertas.Visible = True
        If frmG_Alertas.WindowState = 1 Then frmG_Alertas.WindowState = 0
End Sub

Private Sub mnuBuscaActualizar_Click()
On Error Resume Next
Call frmBuscandoActualización.Show
End Sub

Private Sub mnuCargarDesdeBackUp_Click()
On Error Resume Next
' Barra de progreso!!
FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"


#If UsarAPI Then
Call apiclosesocket(SockListen)
#Else
frmGeneral.Socket1.Cleanup
frmGeneral.Socket2(0).Cleanup
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call CargarBackUp
Call LoadOBJData_Nuevo

#If UsarAPI Then
SockListen = ListenForConnect(Puerto, frmGeneral.hwnd, "")

#Else
frmGeneral.Socket1.AddressFamily = AF_INET
frmGeneral.Socket1.protocol = IPPROTO_IP
frmGeneral.Socket1.SocketType = SOCK_STREAM
frmGeneral.Socket1.Binary = False
frmGeneral.Socket1.Blocking = False
frmGeneral.Socket1.BufferSize = 1024

frmGeneral.Socket2(0).AddressFamily = AF_INET
frmGeneral.Socket2(0).protocol = IPPROTO_IP
frmGeneral.Socket2(0).SocketType = SOCK_STREAM
frmGeneral.Socket2(0).Blocking = False
frmGeneral.Socket2(0).BufferSize = 2048

'Escucha
frmGeneral.Socket1.LocalPort = Puerto
frmGeneral.Socket1.listen
#End If
End Sub

Private Sub mnuCerrarPop_Click()
Call mnuCerrar_Click
End Sub

Private Sub mnuEscaneadordePjs_Click()
On Error Resume Next
EscaneadorDePJs.Show
End Sub

Private Sub mnuGuardarPersonajesYClanes_Click()
If haciendoBK = True Then
    Call FrmMensajes.msg("Alerta", "No puedes guardar personajes y clanes mientras se esta realizando un backup.")
    Exit Sub
End If

Me.MousePointer = 11
Call GuardarUsuarios
Call SaveGuildsDB
Me.MousePointer = 0
Call FrmMensajes.msg("Nota", "Personajes y Clanes guardados!")
End Sub

Private Sub mnuGuardarPOP_Click()
Call mnuGuardarPersonajesYClanes_Click
End Sub

Private Sub mnuHacerBackPOP_Click()
Call mnuHacerBackUp_Click
End Sub

Private Sub mnuHacerBackUp_Click()
On Error GoTo eh
    If haciendoBK = True Then
        Call FrmMensajes.msg("Alerta", "No puedes hacer un backup mientras ya se esta realizando.")
        Exit Sub
    End If
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    Call FrmMensajes.msg("Nota", "WORLDSAVE OK!!")
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")

End Sub

Private Sub mnuMostrar_Click()
On Error Resume Next
    Me.WindowState = vbNormal
    Me.Visible = True
    QuitarSysTray
End Sub


Private Sub mnuNPCsR_Click()
On Error Resume Next
Dim i As Integer
Call SendData(ToAll, 0, 0, "BKW")
Call SendData(ToAll, 0, 0, "||<Host> Actualizando NPCs..." & FONTTYPE_TALK & ENDC)
Call DescargaNpcsDat
Call CargaNpcsDat
Call SendData(ToAll, 0, 0, "||<Host> Recargando NPCs..." & FONTTYPE_TALK & ENDC)
For i = 1 To LastNPC
    Call ResetNPC(i)
Next
DoEvents
Call SendData(ToAll, 0, 0, "||<Host> Listo" & FONTTYPE_TALK & ENDC)
Call SendData(ToAll, 0, 0, "BKW")
End Sub

Private Sub mnuOcultar_Click()

PonerSysTray

End Sub


Private Sub Auditoria_Timer()
On Error GoTo errhand

Dim k As Integer
For k = 1 To LastUser
    If UserList(k).ConnID <> -1 Then
        DayStats.Segundos = DayStats.Segundos + 1
    End If
Next k

Call PasarSegundo

Static Andando As Boolean
Static Contador As Long
Dim Tmp As Boolean

Contador = Contador + 1

If Contador >= 10 Then
    Contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

Exit Sub

errhand:
Call LogError("Error en Timer Auditoria (sistema de desconexion de 10 segundos).")
Call LogError("Err: " & Err.Number & " - " & Err.Description)

End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute


'Call LogCOSAS("Crash-Test", "Inicia..." & str(MinsRunning), False)
'Call SendData(ToAdmins, 0, 0, "||Comienza el minuto " & MinsRunning & FONTTYPE_INFO)



Dim tName As Long
Dim iuserindex As Integer
Static Minutos As Long
Static MinutosLatsClean As Long
Dim i As Integer

Static MinsSocketReset As Long

Static MinsPjesSave As Long

MinsRunning = MinsRunning + 1
' [GS]
If tDeRepetir.Enabled = False Then tDeRepetir.Enabled = True
' [/GS]
HsMantenimiento = HsMantenimiento - 1
If HsMantenimiento < 10 Then
    If HsMantenimiento = 0 Then
        mnuArchivo.Enabled = False
        mnuVer.Enabled = False
        mnuAcciones.Enabled = False
        mnuActualizar.Enabled = False
        mnuPopupMenu.Enabled = False
        
        ' [GS] Nuevo :D
        Call WriteVar(IniPath & "Server.ini", "SEGURIDAD", "Funcionando", 0)
        ' [/GS]
        
        mnuCerrarCorrectamente.Checked = True
        mnuCerrar.Checked = True
        Call SendData(ToAll, 0, 0, "||<Host-Auto> EL SERVIDOR SE ESTA CERRANDO" & FONTTYPE_TALK & ENDC)
        Call SendData(ToAll, 0, 0, "||<Host-Auto> ADIOS :D" & FONTTYPE_TALK & ENDC)
        DoEvents
        Me.MousePointer = 11
        DoEvents
        Call GuardarUsuarios
        DoEvents
        Call SaveGuildsDB
        DoEvents
        FrmStat.Show
        Call DoBackUp
        DoEvents
        
        Me.MousePointer = 0
        DoEvents
        
        For i = 1 To LastUser
            If (UserList(i).Name <> "") Then
                Call CloseUser(i)
            End If
        Next
        
        Call Shell(App.Path & "\" & App.EXEName & ".exe -mantenimiento", vbNormalFocus)
    
        Dim f
        For Each f In Forms
            Unload f
        Next
        End
        
        Exit Sub
    ElseIf HsMantenimiento = 1 Then
        Call SendData(ToAll, 0, 0, "||<MANTENIMIENTO> Queda 1 minutos para el Mantenimiento." & FONTTYPE_TALK & ENDC)
        Call SendData(ToAll, 0, 0, "!!Queda 1 minutos para el Mantenimiento." & ENDC)
    Else
        Call SendData(ToAll, 0, 0, "||<MANTENIMIENTO> Quedan " & HsMantenimiento & " minutos para el Mantenimiento, vaya tomando las precausiones debidas." & FONTTYPE_TALK & ENDC)
    End If
    
End If

i = HsMantenimiento
Do
    If i < 60 Then Exit Do
    i = i - 60
Loop
mnuMan.Caption = "Mant.: " & (HsMantenimiento - i) / 60 & " hs con " & (i) & " minutos."


If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        Call DayElapsed
        'Dias = Dias + 1
        Horas = 0
    End If
    MinsRunning = 0
End If



    
Minutos = Minutos + 1

' [GS] Sistema de POWA-por-Tiempo y Aventura!
For i = 1 To LastUser

    If UserList(i).GuildInfo.BorroClan = True Then UserList(i).GuildInfo.BorroClan = False

    If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged = True Then
        ' Si el usuario esta online y logeado
        UserList(i).flags.TiempoOnline = UserList(i).flags.TiempoOnline + 1
        ' Le sumo 1 minuto al tiempo ya acumulado
        If UserList(i).flags.TiempoOnline > MaxTiempoOn And (UserList(i).flags.Privilegios < 1 And EsAdmin(i) = False) Then
            ' Si el tiempo del usuario es mayor que el tiempo maximo!
            MaxTiempoOn = UserList(i).flags.TiempoOnline
            MaxTINombre = UserList(i).Name
            Call WriteVar(IniPath & "Estadisticas.ini", "POWA-TO", "Nombre", MaxTINombre)
            Call WriteVar(IniPath & "Estadisticas.ini", "POWA-TO", "TiempoOnline", str(MaxTiempoOn))
            ' Guardo los datos
        End If
        
        ' si esta en aventura!!!
        If UserList(i).flags.AV_Esta = True Then
            UserList(i).flags.AV_Tiempo = UserList(i).flags.AV_Tiempo - 1
            ' Le resto 1 minuto
            If UserList(i).flags.AV_Tiempo = 2 Then
                Call SendData(ToIndex, i, 0, "||Faltan 2 minutos para terminar tu aventura." & FONTTYPE_FIGHT_YO)
                ' Informar a los 2 minutos
            ElseIf UserList(i).flags.AV_Tiempo = 1 Then
                Call SendData(ToIndex, i, 0, "||Falta 1 minuto para terminar tu aventura." & FONTTYPE_FIGHT_YO)
                ' Informar al ultimo minuto
            ElseIf UserList(i).flags.AV_Tiempo <= 0 Then
                ' Se le acabo el tiempo
                UserList(i).flags.AV_Esta = False
                Call SendData(ToIndex, i, 0, "||Tu aventura ha terminado." & FONTTYPE_FIGHT)
                Call WarpUserChar(i, val(ReadField(1, UserList(i).flags.AV_Lugar, 45)), val(ReadField(2, UserList(i).flags.AV_Lugar, 45)), val(ReadField(3, UserList(i).flags.AV_Lugar, 45)), True)
                UserList(i).flags.AV_Tiempo = 0
                ' Listo, termino la aventura perfectamente
            End If
        End If
    End If
Next
' [/GS]

'MinsSocketReset = MinsSocketReset + 1
'for debug purposes
'If MinsSocketReset > 1 Then
'    MinsSocketReset = 0

    For i = 1 To MaxUsers
        'If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
        If UserList(i).flags.UserLogged = True Then
            ' Esta online
            UserList(i).flags.TieneMensaje = False
        End If
    Next i
'    Call ReloadSokcet
'End If



' [GS] Loteria bendita
'If Minutos >= 60 Then
'    Call aClon.VaciarColeccion
'    ' [GS] Sortear la Loteria
'    Dim Num1, Num2 As Integer
'    ' Hace los numeros ganadores
'    If ColaLoteria.Longitud > 0 Then ' Si jugo al menos 1
'        Num1 = Format(RandomNumber("00", "99"), "##")
'        Num2 = Format(RandomNumber("00", "99"), "##")
'        Call SendData(ToAll, 0, 0, "||<LOTERIA> Los resultados de la loteria fueron " & str(Num1) & " y " & str(Num2) & FONTTYPE_VENENO & ENDC)
'        If ColaLoteriaNum.Existe(Num1 & " " & Num2) Then
'            Dim Ganadores As Integer
'            Ganadores = 0
'            ' Primero cuenta cuantos han ganado
'            For i = 1 To ColaLoteriaNum.Longitud
'                If ColaLoteriaNum.VerElemento(i) = Num1 & " " & Num2 Then
'                    If FileExist(App.Path & "\Charfile\" & UCase$(ColaLoteria.VerElemento(i)) & ".chr", vbNormal) Then
'                        Ganadores = Ganadores + 1
'                    ElseIf NameIndex(ColaLoteria.VerElemento(i)) > 0 Then
'                        Ganadores = Ganadores + 1
'                    End If
'                End If
'            Next
'            If Ganadores > 1 Then
'                Call SendData(ToAll, 0, 0, "||<LOTERIA> Se registraron " & Ganadores & " ganadores." & FONTTYPE_VENENO & ENDC)
'            ElseIf Ganadores = 1 Then
'                Call SendData(ToAll, 0, 0, "||<LOTERIA> Se registro un unico ganador." & FONTTYPE_VENENO & ENDC)
'            End If
'
'            ' Divide el premio entre los ganadores
'            If Ganadores = 0 Then Ganadores = 1 ' Evito dividir por 0
'            Pozo_Loteria = CLng(Pozo_Loteria / Ganadores)
'
'            ' Busca nuevamente y entre los premios
'            For i = 1 To ColaLoteriaNum.Longitud
'                If ColaLoteriaNum.VerElemento(i) = Num1 & " " & Num2 Then
'                    tName = NameIndex(ColaLoteria.VerElemento(i))
'                    If tName > 0 Then    ' Esta online?
'                        Call SendData(ToAll, 0, 0, "||<LOTERIA> Felicitaciones " & ColaLoteria.VerElemento(i) & " has ganado el pozo." & FONTTYPE_VENENO & ENDC)
'                        Call SendData(ToIndex, Name, 0, "||<LOTERIA> " & ColaLoteria.VerElemento(i) & " tu premio ha sido depositado en tu cuenta bancaria." & FONTTYPE_VENENO & ENDC)
'                        UserList(tName).Stats.banco = UserList(Name).Stats.banco + Pozo_Loteria
'                    Else                ' Esta offline?
'                        If FileExist(App.Path & "\Charfile\" & UCase$(ColaLoteria.VerElemento(i)) & ".chr", vbNormal) Then
'                            Call SendData(ToAll, 0, 0, "||<LOTERIA> Felicitaciones " & ColaLoteria.VerElemento(i) & " has ganado el pozo, aunque se encuentra desconectado en el momento." & FONTTYPE_VENENO & ENDC)
'                            WriteVar App.Path & "\Charfile\" & ColaLoteria.VerElemento(i) & ".chr", "STATS", "BANCO", val(GetVar(App.Path & "\Charfile\" & ColaLoteria.VerElemento(i) & ".chr", "STATS", "BANCO")) + Pozo_Loteria
'                        End If
'                    End If
'                End If
'            Next
'        Else
'            Call SendData(ToAll, 0, 0, "||<LOTERIA> No se encontraron ganadores. El pozo acumulado es de " & Pozo_Loteria & " monedas de oro." & FONTTYPE_VENENO & ENDC)
'        End If
'        Call SendData(ToAll, 0, 0, "||<LOTERIA> Agradecemos a los " & ColaLoteria.Longitud & " personajes que han participado, gracias y vuelvan a participar." & FONTTYPE_VENENO & ENDC)
'        ' Borra los que jugaron
'        ColaLoteria.Reset
'        ' [/GS]
'    End If
'End If
' [/GS]

If MinutosWs <> 0 Then
    If Minutos >= MinutosWs Then
        If haciendoBK = True Then
            Exit Sub
        End If
        
        DoEvents
        Call DoBackUp
        DoEvents
        
        ' [GS] Hace a todos mas activos :P
        For iuserindex = 1 To MaxUsers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iuserindex).ConnID <> -1 And UserList(iuserindex).flags.UserLogged Then
            'Actualiza el contador de inactividad
            UserList(iuserindex).Counters.IdleCount = 0
        End If
        Next iuserindex
        ' [/GS]
    
        Minutos = 0
    End If
End If

If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

'[Consejeros]
'If MinsPjesSave >= 30 Then
'    MinsPjesSave = 0
'    Call GuardarUsuarios
'Else
'    MinsPjesSave = MinsPjesSave + 1
'End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile(1)
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

'Call SendData(ToAdmins, 0, 0, "||Termina el minuto " & MinsRunning & FONTTYPE_INFO)

Exit Sub

errhandler:

    Call LogError("Error en TimerAutoSave")

End Sub

Private Sub CmdExec_Timer()
On Error GoTo falloC
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
    End If
Next i
Exit Sub
falloC:
    Call LogError("Error en CmdExec_Timer")
End Sub

Private Sub FX_Timer()
'On Error GoTo falloFX ' no crash
On Error Resume Next
Dim MapIndex As Integer
Dim N As Integer
' [GS]
If NumMaps = 0 Then Exit Sub
' [/GS]
For MapIndex = 1 To NumMaps
    Randomize
    If RandomNumber(1, 150) < 12 Then

        If MapInfo(MapIndex).NumUsers > 0 Then

                Select Case MapInfo(MapIndex).Terreno
                   'Bosque
                   Case Bosque
                        N = RandomNumber(1, 100)
                        Select Case MapInfo(MapIndex).Zona
                            Case Campo
                              If Not Lloviendo Then
                                If N < 30 And N >= 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf N < 30 And N < 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf N >= 30 And N <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf N >= 35 And N <= 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf N >= 40 And N <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                            Case Ciudad
                               If Not Lloviendo Then
                                If N < 30 And N >= 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf N < 30 And N < 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf N >= 30 And N <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf N >= 35 And N <= 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf N >= 40 And N <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                        End Select

                End Select

        End If
    End If
Next
Exit Sub

'falloFX:
    'Call LogError("Error en FX_Timer - Err: " & Err.Number & " - " & Err.Description)

End Sub

Private Sub GameTimer_Timer()
On Error GoTo FalloGT ' :D

' [GS] BK?
If haciendoBK = True Then Exit Sub
' [/GS]

Dim iuserindex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer

Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long

'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If

If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If

If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]


 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iuserindex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iuserindex).ConnID <> -1 Then
      '¿User valido?
      If UserList(iuserindex).flags.UserLogged Then
         
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         
         UserList(iuserindex).NumeroPaquetesPorMiliSec = 0

         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>
'         If lPermiteAtacar < IntervaloUserPuedeAtacar Then
'                lPermiteAtacar = lPermiteAtacar + 1
'         Else
         If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
                UserList(iuserindex).flags.PuedeAtacar = 1
'                lPermiteAtacar = 0
         End If
         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>

         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>
'         If lPermiteCast < IntervaloUserPuedeCastear Then
'              lPermiteCast = lPermiteCast + 1
'         Else
         If Not lPermiteCast < IntervaloUserPuedeCastear Then
                UserList(iuserindex).flags.PuedeLanzarSpell = 1
                If UserList(iuserindex).flags.QuiereLanzarSpell = 1 Then
                    UserList(iuserindex).flags.QuiereLanzarSpell = 0
                    Call SendData(ToIndex, iuserindex, 0, "T01" & Magia)
                End If
'              lPermiteCast = 0
         End If
         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>

         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>
         If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
              lPermiteTrabajar = lPermiteTrabajar + 1
         ElseIf Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
              UserList(iuserindex).flags.PuedeTrabajar = 1
              lPermiteTrabajar = 0
         End If
         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>


         Call DoTileEvents(iuserindex, UserList(iuserindex).Pos.Map, UserList(iuserindex).Pos.X, UserList(iuserindex).Pos.Y)
         
                
         If UserList(iuserindex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iuserindex)
         If UserList(iuserindex).flags.Ceguera = 1 Or _
            UserList(iuserindex).flags.Estupidez Then Call EfectoCegueEstu(iuserindex)
          
         If UserList(iuserindex).flags.Muerto = 0 Then
               
               '[Consejeros]
               If UserList(iuserindex).flags.Desnudo And (UserList(iuserindex).flags.Privilegios = 0 And EsAdmin(iuserindex) = False) Then Call EfectoFrio(iuserindex)
               If UserList(iuserindex).flags.Meditando Then Call DoMeditar(iuserindex)
               If UserList(iuserindex).flags.Envenenado = 1 And (UserList(iuserindex).flags.Privilegios = 0 And EsAdmin(iuserindex) = False) Then Call EfectoVeneno(iuserindex, bEnviarStats)
               If UserList(iuserindex).flags.AdminInvisible <> 1 And UserList(iuserindex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iuserindex)
               ' v0.12a9
               If UserList(iuserindex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iuserindex)
          
               Call DuracionPociones(iuserindex)
               Call HambreYSed(iuserindex, bEnviarAyS)
               ' [GS] Duracion Efecto Magico de Terreno
               Call DuracionExplocionMagica(iuserindex)
               ' [/GS]
               
               ' 0.12b3
               If (UserList(iuserindex).flags.Desnudo = 0 And BajaStamina = True) Or BajaStamina = False Then

               If Lloviendo Then
                    If Not Intemperie(iuserindex) Then
                                 If Not UserList(iuserindex).flags.Descansar And (UserList(iuserindex).flags.Hambre = 0 And UserList(iuserindex).flags.Sed = 0) Then
                                 'No esta descansando
                                          Call Sanar(iuserindex, bEnviarStats, SanaIntervaloSinDescansar)
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                 ElseIf UserList(iuserindex).flags.Descansar Then
                                 'esta descansando
                                          Call Sanar(iuserindex, bEnviarStats, SanaIntervaloDescansar)
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                                          'termina de descansar automaticamente
                                          If UserList(iuserindex).Stats.MaxHP = UserList(iuserindex).Stats.MinHP And _
                                             UserList(iuserindex).Stats.MaxSta = UserList(iuserindex).Stats.MinSta Then
                                                    Call SendData(ToIndex, iuserindex, 0, "DOK")
                                                    Call SendData(ToIndex, iuserindex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                                    UserList(iuserindex).flags.Descansar = False
                                          End If
                                 End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iuserindex).flags.Descansar And (UserList(iuserindex).flags.Hambre = 0 And UserList(iuserindex).flags.Sed = 0) Then
                    'No esta descansando
                             Call Sanar(iuserindex, bEnviarStats, SanaIntervaloSinDescansar)
                             Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                    ElseIf UserList(iuserindex).flags.Descansar Then
                    'esta descansando
                             Call Sanar(iuserindex, bEnviarStats, SanaIntervaloDescansar)
                             Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                             'termina de descansar automaticamente
                             If UserList(iuserindex).Stats.MaxHP = UserList(iuserindex).Stats.MinHP And _
                                UserList(iuserindex).Stats.MaxSta = UserList(iuserindex).Stats.MinSta Then
                                     Call SendData(ToIndex, iuserindex, 0, "DOK")
                                     Call SendData(ToIndex, iuserindex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                     UserList(iuserindex).flags.Descansar = False
                             End If
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
               
               End If

               If bEnviarStats Then Call SendUserStatsBox(iuserindex)
               If bEnviarAyS Then Call EnviarHambreYsed(iuserindex)

               If UserList(iuserindex).NroMacotas > 0 Then Call TiempoInvocacion(iuserindex)
       End If 'Muerto
     Else 'no esta logeado?
     'UserList(iUserIndex).Counters.IdleCount = 0
     '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
     'de dados :)
        UserList(iuserindex).Counters.IdleCount = UserList(iuserindex).Counters.IdleCount + 1
        If UserList(iuserindex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iuserindex).Counters.IdleCount = 0
              Call CloseSocket(iuserindex)
        End If
     End If 'UserLogged
        
   End If

  Next iuserindex
   
'[Alejo]
If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If

If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If

If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If

'[/Alejo]
  'DoEvents
Exit Sub

FalloGT:
    Call LogError("Error en GameTimer_Timer - UserIndex: " & iuserindex & " Err: " & Err.Number & " - " & Err.Description)

End Sub

Private Sub KillLog_Timer()
On Error Resume Next

If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"

If frmGeneral.Visible = False And frmCargando.Visible = False Then
    Call PonerSysTray
End If

End Sub

Private Sub MDIForm_Click()
On Error Resume Next
frmG_Main.Show
frmG_Main.SetFocus
If frmG_Main.WindowState = 1 Then frmG_Main.WindowState = 0
    
End Sub





'Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    'Dim nidTemp As NOTIFYICONDATA

    'nidTemp.cbSize = Len(nidTemp)
    'nidTemp.hwnd = hwnd
    'nidTemp.uId = ID
    'nidTemp.uFlags = flags
    'nidTemp.uCallbackMessage = CallbackMessage
    'nidTemp.hIcon = Icon
    'nidTemp.szTip = Tip & Chr$(0)

'    setNOTIFYICONDATA = nidTemp
'End Function



'Private Sub QuitarIconoSystray()
'On Error Resume Next

'Borramos el icono del systray
'Dim i As Integer
'Dim nid As NOTIFYICONDATA

'nid = setNOTIFYICONDATA(frmGeneral.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmGeneral.Icon, "")

'i = Shell_NotifyIconA(NIM_DELETE, nid)
'

'End Sub



Sub CheckIdleUser()
On Error GoTo FalloCIU
Dim iuserindex As Integer

For iuserindex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
   If UserList(iuserindex).ConnID <> -1 And UserList(iuserindex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        ' [GS] Sistema antilogeo
        If UserList(iuserindex).flags.RecienIni = True Then
            UserList(iuserindex).flags.TiempoIni = UserList(iuserindex).flags.TiempoIni + 1
            If UserList(iuserindex).flags.TiempoIni > 1 Then
                ' Esta logeado!!!
                If UserList(iuserindex).Pos.Map = MapaAventura Then
                    UserList(iuserindex).flags.AV_Esta = False
                    UserList(iuserindex).flags.AV_Tiempo = 0
                End If
                If UserList(iuserindex).Pos.Map = Ullathorpe.Map Then
                    ' nix
                    Call WarpUserChar(iuserindex, Nix.Map, Nix.X, Nix.Y, True)
                ElseIf UserList(iuserindex).Pos.Map = Nix.Map Then
                    ' bander
                    Call WarpUserChar(iuserindex, Banderbill.Map, Banderbill.X, Banderbill.Y, True)
                ElseIf UserList(iuserindex).Pos.Map = Banderbill.Map Then
                    ' mover a lindos
                    Call WarpUserChar(iuserindex, Lindos.Map, Lindos.X, Lindos.Y, True)
                ElseIf UserList(iuserindex).Pos.Map = Lindos.Map Then
                    ' mover a mapa 2
                    Call WarpUserChar(iuserindex, 2, 10, 15, True)
                Else
                    ' mover a ulla
                    Call WarpUserChar(iuserindex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
                End If
                Call SendData(ToIndex, iuserindex, 0, "||Has sido deslogeado por el sistema automatico." & FONTTYPE_FIGHT_YO)
                UserList(iuserindex).flags.TiempoIni = 0
            End If
        Else
        ' [/GS]
            UserList(iuserindex).Counters.IdleCount = UserList(iuserindex).Counters.IdleCount + 1
            If UserList(iuserindex).Counters.IdleCount >= IdleLimit Then
                If haciendoBK = True Then
                    UserList(iuserindex).Counters.IdleCount = 0
                Else
                    Call SendData(ToIndex, iuserindex, 0, "!!Demasiado tiempo inactivo.")
                    Call Cerrar_Usuario(iuserindex)
                End If
            End If
        End If
    End If
Next iuserindex
Exit Sub
FalloCIU:
    Call LogError("Error en CheckIdleUser - Err: " & Err.Number & " - " & Err.Description)

End Sub


Public Sub InitMain(ByVal f As Byte)
On Error Resume Next
If f = 1 Then
    Call mnuOcultar_Click
    DoEvents
Else
    frmGeneral.Show
End If

End Sub


Private Sub mnuAcerca_Click()
On Error Resume Next
Call frmCreditos.Show
'Call FrmMensajes.msg("Creditos", "GS Server AO " & Me.Tag & vbCrLf & "Programado por ^[GS]^" & vbCrLf & "Web site: http://www.gs-zone.com.ar" & vbCrLf & "E-mail: gshaxor@gmail.com" & vbCrLf & "(r) NMS Optimized" & vbCrLf & vbCrLf & "Agradecimientos: Ver /CREDITOS, en el juego. ;)")
End Sub

Private Sub mnuCerrar_Click()
On Error Resume Next
If haciendoBK = True Then
    Call FrmMensajes.msg("Alerta", "No puedes cerrar mientras realizas un backup.")
    Exit Sub
End If
' [GS] Nuevo :D
Call WriteVar(IniPath & "Server.ini", "SEGURIDAD", "Funcionando", 0)
' [/GS]
Call GuardarUsuarios
DoEvents
Call SaveGuildsDB
DoEvents
If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    mnuCerrar.Checked = True
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnuCerrarCorrectamente_Click()
On Error Resume Next
DoEvents
If haciendoBK = True Then
    Call FrmMensajes.msg("Alerta", "No puedes cerrar mientras realizas un backup.")
    Exit Sub
End If
If MsgBox("¿Desea cerrar el servidor correctamente?" & vbCrLf & "NOTA: Puede tardar unos minutos.", vbYesNo) = vbYes Then

    mnuArchivo.Enabled = False
    mnuVer.Enabled = False
    mnuAcciones.Enabled = False
    mnuActualizar.Enabled = False
    mnuPopupMenu.Enabled = False
    
    ' [GS] Nuevo :D
    Call WriteVar(IniPath & "Server.ini", "SEGURIDAD", "Funcionando", 0)
    ' [/GS]
    
    mnuCerrarCorrectamente.Checked = True
    mnuCerrar.Checked = True
    Call SendData(ToAll, 0, 0, "||<Host-Auto> EL SERVIDOR SE ESTA CERRANDO" & FONTTYPE_TALK & ENDC)
    Call SendData(ToAll, 0, 0, "||<Host-Auto> ADIOS :D" & FONTTYPE_TALK & ENDC)
    DoEvents
    Me.MousePointer = 11
    DoEvents
    Call GuardarUsuarios
    DoEvents
    Call SaveGuildsDB
    DoEvents
    FrmStat.Show
    Call DoBackUp
    DoEvents

    Me.MousePointer = 0
    DoEvents
    Dim f
    For Each f In Forms
        Unload f
    Next
End If
DoEvents
End Sub

Private Sub mnuConfiguración_Click()
frmG_Configurar.Show
End Sub

Private Sub mnuGuardiasPos_Click()
Call SendData(ToAll, 0, 0, "BKW")
Call SendData(ToAll, 0, 0, "||<Host> Actualizando posicion de los guardias..." & FONTTYPE_TALK & ENDC)
Call ReSpawnOrigPosNpcs
Call SendData(ToAll, 0, 0, "||<Host> Listo" & FONTTYPE_TALK & ENDC)
Call SendData(ToAll, 0, 0, "BKW")

End Sub

Private Sub mnuHechizos_Click()
Call SendData(ToAll, 0, 0, "BKW")
Call SendData(ToAll, 0, 0, "||<Host> Actualizando Hechizos..." & FONTTYPE_TALK & ENDC)
Call CargarHechizos
Call SendData(ToAll, 0, 0, "||<Host> Listo" & FONTTYPE_TALK & ENDC)
Call SendData(ToAll, 0, 0, "BKW")

End Sub




Private Sub mnuMOTD_Click()
Call LoadMotd
End Sub

Private Sub mnuNombresProhibidos_Click()
Call CargarForbidenWords
End Sub

Private Sub mnuNPCDebug_Click()
frmG_NPCDebug.Show
End Sub

Private Sub mnuNPCs_Click()
Call SendData(ToAll, 0, 0, "BKW")
Call SendData(ToAll, 0, 0, "||<Host> Actualizando NPCs..." & FONTTYPE_TALK & ENDC)
Call DescargaNpcsDat
Call CargaNpcsDat
Call SendData(ToAll, 0, 0, "||<Host> Listo" & FONTTYPE_TALK & ENDC)
Call SendData(ToAll, 0, 0, "BKW")
End Sub

Private Sub mnuObjetos_Click()
Call SendData(ToAll, 0, 0, "BKW")
Call SendData(ToAll, 0, 0, "||<Host> Actualizando Objetos..." & FONTTYPE_TALK & ENDC)
Call LoadOBJData_Nuevo
Call SendData(ToAll, 0, 0, "||<Host> Listo" & FONTTYPE_TALK & ENDC)
Call SendData(ToAll, 0, 0, "BKW")

End Sub


Private Sub mnuPaneldeControl_Click()
frmG_Main.Show
End Sub

Private Sub mnuPOPH_Click()
Call mnuHechizos_Click
End Sub

Private Sub mnuPOPN_Click()
Call mnuNPCs_Click
End Sub

Private Sub mnuPOPNR_Click()
Call mnuNPCsR_Click
End Sub

Private Sub mnuPOPP_Click()
Call mnuGuardiasPos_Click
End Sub

Private Sub mnuPOPUNBANIP_Click()
Call mnuUnBanIP_Click
End Sub

Private Sub mnuReloadOpciones_Click()
Call LoadOpcsINI
End Sub

Private Sub mnuReLoadSpawn_Click()
Call CargarSpawnList
End Sub

Private Sub mnuServidorINI_Click()
Call LoadSini
'Call frmG_Main.cmdResetear_Click
Call FrmMensajes.msg("Nota", "Server.ini recargado.")
End Sub

Private Sub mnuSocketDebug_Click()
frmG_Sockets.Show
End Sub

Private Sub mnuTrafico_Click()
frmG_Trafico.Show
End Sub

Private Sub mnuUnBAN_Click()
On Error Resume Next

Dim Fn As String
Dim cad$
Dim N As Integer, k As Integer

Fn = App.Path & "\logs\GenteBanned.log"

If FileExist(Fn, vbNormal) Then
    N = FreeFile
    Open Fn For Input Shared As #N
    Do While Not EOF(N)
        k = k + 1
        Input #N, cad$
        Call UnBan(cad$)
        
    Loop
    Close #N
    Call FrmMensajes.msg("Nota", "Se han habilitado " & k & " personajes.")
    Kill Fn
End If

End Sub

Private Sub mnuUnBanIP_Click()
Dim i As Long, N As Long

N = BanIps.Count
For i = 1 To BanIps.Count
    BanIps.Remove 1
Next i

Call FrmMensajes.msg("Nota", "Se han habilitado " & N & " IP's")

End Sub

Private Sub mnuUsuarios_Click()
End Sub


Private Sub mnuValidClien_Click()
frmG_ValCliente.Show
End Sub

Private Sub mnuWEB_Click()
On Error Resume Next
Call Shell("explorer http://www.gs-zone.com.ar", vbMaximizedFocus)
End Sub

Private Sub NoMain_Timer()
On Error GoTo fallo
If frmCargando.Visible = True And Me.Visible = True Then
    Me.Hide
End If

Exit Sub
fallo:
End Sub

Private Sub NpcAtaca_Timer()
On Error GoTo FalloNA
Dim Npc As Integer

For Npc = 1 To LastNPC
    Npclist(Npc).CanAttack = 1
Next Npc

Exit Sub
FalloNA:
    Call LogError("Error en NpcAtaca_Timer")

End Sub

Private Sub sendMSGx_Click()
On Error Resume Next
If Mensaje <> "" Then
    If Por.ListIndex <> -1 And Dirigido.ListIndex <> -1 Then
        If Por.ListIndex = 0 Then ' Consola
            If Dirigido.ListIndex = 0 Then ' TODOS
                Call SendData(ToAll, 0, 0, "||<Host>" & Mensaje.Text & FONTTYPE_TALK & ENDC)
            Else ' GMs
                Call SendData(ToAdmins, 0, 0, "||" & "GM's: " & Mensaje.Text & FONTTYPE_FIGHT & ENDC)
            End If
            Mensaje.Text = ""
        Else    ' En ventana
            If Dirigido.ListIndex = 0 Then ' TODOS
                Call SendData(ToAll, 0, 0, "!!" & Mensaje.Text & ENDC)
            Else ' GMs
                Call SendData(ToAdmins, 0, 0, "!!" & Mensaje.Text & ENDC)
            End If
            Mensaje.Text = ""
        End If
    End If
End If
End Sub


Private Sub Master_ConnectionRequest(ByVal requestID As Long)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
On Error Resume Next
    Dim i As Long
    
    For i = 0 To 200
        If Slave(i).State = sckClosed Then
            Slave(i).Close
            Slave(i).accept requestID
            Exit Sub
        End If
    Next i
End Sub

Private Sub Slave_DataArrival(Index As Integer, ByVal BytesTotal As Long)
On Error GoTo fallo
    Dim strData As String
    Dim strGet As String
    Dim spc2 As Long
    Dim page As String
   
    Slave(Index).GetData strData
    
    strData = ConvertUTF8toASCII(strData)
   
    'For the Get server command
    If Mid(strData, 1, 3) = "GET" Then
        strGet = InStr(strData, "GET ")
        spc2 = InStr(strGet + 5, strData, " ")
        page = Trim(Mid(strData, strGet + 5, spc2 - (strGet + 4)))
        If Right(page, 1) = "/" Then page = Left(page, Len(page) - 1)
        If page = "/" Then page = "index.html"
        If page = "" Then page = "index.html"
        If page = "index.html" Then
            Slave(Index).SendData IndexData
            Exit Sub
        Else
            Slave(Index).SendData DarData(page)
            Exit Sub
        End If
        Slave(Index).Close
    End If
Exit Sub
fallo:
Slave(Index).SendData "<html>ERROR INTERNO</html>"
DoEvents
End Sub

Private Sub Slave_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
    Slave(Index).Close
End Sub

Private Sub Slave_SendComplete(Index As Integer)
On Error Resume Next
    Slave(Index).Close
End Sub

Private Sub tDeRepetir_Timer()
On Error Resume Next
If UPdata < 0 Then UPdata = 0
If DLdata < 0 Then DLdata = 0
Dim Temp1 As String
Dim Temp2 As String
Dim Temp3 As String
Dim Temp4 As String
Temp1 = IIf(IsNumeric(Left(Format(UPdata / 1024, "###.#"), 1)), Format(UPdata / 1024, "###.#"), "0" & Format(UPdata / 1024, "###.#"))
Temp2 = IIf(IsNumeric(Left(Format((UPdata * 8) / 1024, "###.#"), 1)), Format((UPdata * 8) / 1024, "###.#"), "0" & Format((UPdata * 8) / 1024, "###.#"))
Temp3 = IIf(IsNumeric(Left(Format(DLdata / 1024, "###.#"), 1)), Format(DLdata / 1024, "###.#"), "0" & Format(DLdata / 1024, "###.#"))
Temp4 = IIf(IsNumeric(Left(Format((DLdata * 8) / 1024, "###.#"), 1)), Format((DLdata * 8) / 1024, "###.#"), "0" & Format((DLdata * 8) / 1024, "###.#"))
If Right(Temp1, 1) = "," Then Temp1 = Temp1 & "0"
If Right(Temp2, 1) = "," Then Temp2 = Temp2 & "0"
If Right(Temp3, 1) = "," Then Temp3 = Temp3 & "0"
If Right(Temp4, 1) = "," Then Temp4 = Temp4 & "0"

Me.Caption = "Argentum Online - GS Server AO " & Me.Tag & " - Up: " _
    & Temp1 & " Kb/s - [" _
    & Temp2 & " Kbps] - Down: " _
    & Temp3 & " Kb/s - [" _
    & Temp4 & " Kbps]"
UPdata = 0
DLdata = 0
' [GS] Se asegura que se lea el lag
If frmGeneral.tDeRepetir.Enabled = False Then frmGeneral.tDeRepetir.Enabled = True
' [/GS]
'Dim h As Long
'For h = 1 To LastUser
'    If UserList(h).ConnID <> -1 Then
        ' ahora todos pueden volver a resucitarse
'        If UserList(h).flags.Resucitar = False Then
'            UserList(h).flags.Resucitar = True
        ' Ahora todos pueden volver a curarse
'        ElseIf UserList(h).flags.Curacion = False Then
'            UserList(h).flags.Curacion = True
'        End If
'    End If
'Next h

End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler

Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer ' [EL OSO]

If Not haciendoBK Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            ' [EL OSO]
            e_p = esPretoriano(NpcIndex)
            If (e_p > 0) Then
                If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then Call EfectoParalisisNpc(NpcIndex)
                ''''''''''''''''''
                Select Case e_p
                    Case 1  ''clerigo
                        Call PRCLER_AI(NpcIndex)
                    Case 2  ''mago
                        Call PRMAGO_AI(NpcIndex)
                    Case 3  ''cazador
                        Call PRCAZA_AI(NpcIndex)
                    Case 4  ''rey
                        Call PRREY_AI(NpcIndex)
                    Case 5
                        Call PRGUER_AI(NpcIndex)
                End Select
                ''''''''''''''''''
            Else
                ' [/EL OSO]
                ''IA comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> ESTATICO Then
                                    If Npclist(NpcIndex).Meditando = True Then
                                        Call NPCMeditar$(NpcIndex)
                                    Else
                                        Call NPCAI$(NpcIndex)
                                    End If
                                  End If
                          End If
                     End If
                     
                End If
            End If
        End If
    
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 If Npclist(NpcIndex).Pos.Map <> 0 Then Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
    If UserList(i).ConnID = -1 And UserList(i).flags.UserLogged Then
        Call SendData(ToIndex, i, 0, "FINOK")
        Call CloseUser(i)
    End If
Next i

End Sub

Private Sub TimerCartelito_Timer()
On Error Resume Next
If FrmStat.Visible = True And haciendoBK = False Then FrmStat.Visible = False
If frmCargando.Visible = True Then Exit Sub
'If haciendoBK = True Then
'    XEstadoX 3
'Else
'    XEstadoX 1
'End If

End Sub

Private Sub tLluvia_Timer()
On Error GoTo errhandler

Dim iCount As Integer

If Lloviendo Then
   For iCount = 1 To LastUser
    Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
errhandler:
Call LogError("tLluvia")
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler

Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    ' DESACTIVA LA LLUVIA MOLESTA
    If LluviaON = True Then
        MinutosSinLluvia = MinutosSinLluvia + 1
        If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
                If RandomNumber(1, 100) <= 10 Then
                    Lloviendo = True
                    MinutosSinLluvia = 0
                    Call SendData(ToAll, 0, 0, "LLU")
                End If
        ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(ToAll, 0, 0, "LLU")
        End If
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(ToAll, 0, 0, "LLU")
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 7 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(ToAll, 0, 0, "LLU")
            End If
    End If
End If


Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
' Boton elegido
Select Case Button.Index
    Case 1  ' Usuarios
        frmG_Main.Show
        frmG_Main.SetFocus
        If frmG_Main.WindowState = 1 Then frmG_Main.WindowState = 0
    Case 2  ' NPC's Debug
        frmG_NPCDebug.Show
        frmG_NPCDebug.SetFocus
        If frmG_NPCDebug.WindowState = 1 Then frmG_NPCDebug.WindowState = 0
    Case 3  ' Sockets
        frmG_Sockets.Show
        frmG_Sockets.SetFocus
        If frmG_Sockets.WindowState = 1 Then frmG_Sockets.WindowState = 0
    Case 4  ' Trafico
        frmG_Trafico.Show
        frmG_Trafico.SetFocus
        If frmG_Trafico.WindowState = 1 Then frmG_Trafico.WindowState = 0
    Case 5  ' Configuracion
        frmG_Configurar.Show
        frmG_Configurar.SetFocus
        If frmG_Configurar.WindowState = 1 Then frmG_Configurar.WindowState = 0
    Case Else  ' Estado?
        frmG_Alertas.Show
        frmG_Alertas.Visible = True
        If frmG_Alertas.WindowState = 1 Then frmG_Alertas.WindowState = 0
End Select
End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
' [GS] BK?
If haciendoBK = True Then Exit Sub
' [/GS]
' [GS] Se asegura que se lea el lag
If frmGeneral.tDeRepetir.Enabled = False Then frmGeneral.tDeRepetir.Enabled = True
' [/GS]
'Call SendData(ToAdmins, 0, 0, "||Comienza el anti-piquete." & FONTTYPE_INFO)

Static Segundos As Integer

Segundos = Segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
            If UserList(i).flags.Paralizado = 1 Then Exit Sub
            
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = 5 Then
                ' Si es un sendero
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                ' Suma 1 sec al user
                    Call SendData(ToIndex, i, 0, "||Estas obstruyendo la via publica, muevete o seras encarcelado!!!" & FONTTYPE_INFO)
                ' Le dice que esa obstruyendo
                    If UserList(i).Counters.PiqueteC > 23 Then
                        ' Si paso mas de 23 segundos, lo manda a la carcel
                            UserList(i).Counters.PiqueteC = 0
                            If UserList(i).flags.Privilegios > 0 Or EsAdmin(i) = True Then Exit Sub
                            Call Encarcelar(i, 3)
                    End If
            ' [GS] Parado sobre objeto?
            ElseIf UserList(i).flags.Muerto = 1 And MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).OBJInfo.ObjIndex > 0 Then
                ' Esta muerto el user y parado sobre un objeto?
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                ' Le sumo 1 sec
                Call SendData(ToIndex, i, 0, "||Estas obstruyendo un objeto, si no te mueves seras desconectado en 18 segundos!!!" & FONTTYPE_INFO)
                ' Le informo
                If UserList(i).Counters.PiqueteC > 16 Then
                    UserList(i).Counters.PiqueteC = 0
                    Call SendData(ToIndex, i, 0, "||Has sido desconectado por obstruir un objeto!!" & FONTTYPE_INFO)
                    Call CloseSocket(i)
                End If
            ' [/GS]
            Else
                    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If
            
            If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
            
    End If
Next i

If Segundos >= 18 Then Segundos = 0
   
'Call SendData(ToAdmins, 0, 0, "||Termina el anti-piquete." & FONTTYPE_INFO)
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer")

End Sub

Private Sub tTraficStat_Timer()


'Dim i As Integer
'
'If frmTrafic.Visible Then frmTrafic.lstTrafico.Clear
'
'For i = 1 To LastUser
'    If UserList(i).Flags.UserLogged Then
'        If frmTrafic.Visible Then
'            frmTrafic.lstTrafico.AddItem UserList(i).Name & " " & UserList(i).BytesTransmitidosUser + UserList(i).BytesTransmitidosSvr & " bytes per second"
'        End If
'        UserList(i).BytesTransmitidosUser = 0
'        UserList(i).BytesTransmitidosSvr = 0
'    End If
'Next i


End Sub



Private Sub Socket1_Accept(SocketId As Integer)
#If Not (UsarAPI = 1) Then

'=========================================================
'USO DEL CONTROL SOCKET WRENCH
'=============================

If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Pedido de conexion SocketID:" & SocketId & vbCrLf

On Error Resume Next
    
    Dim NewIndex As Integer
    
    
    If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "NextOpenUser" & vbCrLf
    
    NewIndex = NextOpenUser ' Nuevo indice
    If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "UserIndex asignado " & NewIndex & vbCrLf
    
    If NewIndex <= MaxUsers Then
            If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Cargando Socket " & NewIndex & vbCrLf
            
            Unload Socket2(NewIndex)
            Load Socket2(NewIndex)
            
            Socket2(NewIndex).AddressFamily = AF_INET
            Socket2(NewIndex).protocol = IPPROTO_IP
            Socket2(NewIndex).SocketType = SOCK_STREAM
            Socket2(NewIndex).Binary = False
            Socket2(NewIndex).BufferSize = SOCKET_BUFFER_SIZE
            Socket2(NewIndex).Blocking = False
            Socket2(NewIndex).Linger = 1
            
            Socket2(NewIndex).accept = SocketId
            
            Call aDos.Corregir(Socket2(NewIndex).PeerAddress)
            
            If aDos.MaxConexiones(Socket2(NewIndex).PeerAddress) Then
                            
                If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "ERROR: User " & NewIndex & " a llegado al maximo de conecciones." & vbCrLf
                
                UserList(NewIndex).ConnID = -1
                
                If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "User Slot Reseteado " & NewIndex & vbCrLf
                
                If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Socket " & NewIndex & " cerrado." & vbCrLf
                
                'Call LogCriticEvent(Socket2(NewIndex).PeerAddress & " intento crear mas de 3 conexiones.")
                Call aDos.RestarConexion(Socket2(NewIndex).PeerAddress)
                'Socket2(NewIndex).Disconnect
                Unload frmGeneral.Socket2(NewIndex)
                
                Exit Sub
            End If
            
            UserList(NewIndex).ConnID = SocketId
            UserList(NewIndex).IP = Socket2(NewIndex).PeerAddress
            
            If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & Socket2(NewIndex).PeerAddress & " logged." & vbCrLf
    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
    End If
    
Exit Sub

#End If
End Sub


Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
' solo para depurar
'Call LogError("Socket1:" & ErrorString)
On Error Resume Next
If DebugSocket Then frmG_Sockets.Errores.Text = frmG_Sockets.Errores.Text & Time & " " & ErrorString & vbCrLf
'XEstadoX 4
frmG_Sockets.Estado = "Error Socket 1 - " & str(Socket1.State)
End Sub

Private Sub Socket2_Blocking(Index As Integer, Status As Integer, Cancel As Integer)
'Cancel = True
End Sub

Private Sub Socket2_Connect(Index As Integer)
'If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Conectado" & vbCrLf
Set UserList(Index).CommandsBuffer = New CColaArray
End Sub

Private Sub Socket2_Disconnect(Index As Integer)
    If UserList(Index).flags.UserLogged And _
        UserList(Index).Counters.Saliendo = False Then
        Call Cerrar_Usuario(Index)
    Else
        Call CloseSocket(Index)
    End If
End Sub

'Private Sub Socket2_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
''24004   WSAEINTR    Blocking function was canceled
''24009   WSAEBADF    Invalid socket descriptor passed to function
''24013   WSAEACCES   Access denied
''24014   WSAEFAULT   Invalid address passed to function
''24022   WSAEINVAL   Invalid socket function call
''24024   WSAEMFILE   No socket descriptors are available
''24035   WSAEWOULDBLOCK  Socket would block on this operation
''24036   WSAEINPROGRESS  Blocking function in progress
''24037   WSAEALREADY Function being canceled has already completed
''24038   WSAENOTSOCK Invalid socket descriptor passed to function
''24039   WSAEDESTADDRREQ Destination address is required
''24040   WSAEMSGSIZE Datagram was too large to fit in specified buffer
''24041   WSAEPROTOTYPE   Specified protocol is the wrong type for this socket
''24042   WSAENOPROTOOPT  Socket option is unknown or unsupported
''24043   WSAEPROTONOSUPPORT  Specified protocol is not supported
''24044   WSAESOCKTNOSUPPORT  Specified socket type is not supported in this address family
''24045   WSAEOPNOTSUPP   Socket operation is not supported
''24046   WSAEPFNOSUPPORT Specified protocol family is not supported
''24047   WSAEAFNOSUPPORT Specified address family is not supported by this protocol
''24048   WSAEADDRINUSE   Specified address is already in use
''24049   WSAEADDRNOTAVAIL    Specified address is not available
''24050   WSAENETDOWN Network subsystem has failed
''24051   WSAENETUNREACH  Network cannot be reached from this host
''24052   WSAENETRESET    Network dropped connection on reset
''24053   WSAECONNABORTED Connection was aborted due to timeout or other failure
''24054   WSAECONNRESET   Connection was reset by remote network
''24055   WSAENOBUFS  No buffer space is available
''24056   WSAEISCONN  Socket is already connected
''24057   WSAENOTCONN Socket Is Not Connected
''24058   WSAESHUTDOWN    Socket connection has been shut down
''24060   WSAETIMEDOUT    Operation timed out before completion
''24061   WSAECONNREFUSED Connection refused by remote network
''24064   WSAEHOSTDOWN    Remote host is down
''24065   WSAEHOSTUNREACH Remote host is unreachable
''24091   WSASYSNOTREADY  Network subsystem is not ready for communication
''24092   WSAVERNOTSUPPORTED  Requested version is not available
''24093   WSANOTINITIALIZED   Windows sockets library not initialized
''25001   WSAHOST_NOT_FOUND   Authoritative Answer Host not found
''25002   WSATRY_AGAIN    Non-authoritative Answer Host not found
''25003   WSANO_RECOVERY  Non-recoverable error
''25004   WSANO_DATA  No data record of requested type
''Response = SOCKET_ERRIGNORE
'If ErrorCode = 24053 Then Call CloseSocket(Index)
'End Sub


Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)

#If Not (UsarAPI = 1) Then

On Error GoTo ErrorHandler

'*********************************************
'Separamos las lineas con ENDC y las enviamos a HandleData()
'*********************************************
Dim LoopC As Integer
Dim RD As String
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer
Dim aux$
Dim OrigCad As String

Dim LenRD As Long

'<<<<<<<<<<<<<<<<<< Evitamos DoS >>>>>>>>>>>>>>>>>>>>>>>>>>>
Call AddtoVar(UserList(Index).NumeroPaquetesPorMiliSec, 1, 1000)
'
If UserList(Index).NumeroPaquetesPorMiliSec > 700 Then
'   'UserList(Index).Flags.AdministrativeBan = 1
   Call LogCriticalHackAttemp(UserList(Index).Name & " " & frmGeneral.Socket2(Index).PeerAddress & " alcanzo el max paquetes por iteracion.")
   Call SendData(ToIndex, Index, 0, "ERRSe ha perdido la conexion, por favor vuelva a conectarse.")
   Call CloseSocket(Index)
   Exit Sub
End If

Call Socket2(Index).Read(RD, DataLength)

OrigCad = RD
LenRD = Len(RD)
' [GS] Download Data!!

DLdata = DLdata + LenRD

' [/GS]
'Call AddtoVar(UserList(Index).BytesTransmitidosUser, LenB(RD), 100000)

'[¡¡BUCLE INFINITO!!]'
If LenRD = 0 Then
    UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
    If UserList(Index).AntiCuelgue >= 150 Then
        UserList(Index).AntiCuelgue = 0
        'Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
        Socket2(Index).Disconnect
        Call CloseSocket(Index)
        Exit Sub
    End If
Else
    UserList(Index).AntiCuelgue = 0
End If
'[¡¡BUCLE INFINITO!!]'

'Verificamos por una comando roto y le agregamos el resto
If UserList(Index).RDBuffer <> "" Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

'Verifica por mas de una linea
sChar = 1
For LoopC = 1 To LenRD

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

'Verifica una linea rota y guarda
If Len(RD) - (sChar - 1) <> 0 Then
    UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))
End If

'Enviamos el buffer al manejador
For LoopC = 1 To CR
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
    '%%% EL PROBLEMA DEL SPEEDHACK          %%%
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If ClientsCommandsQueue = 1 Then
        If rBuffer(LoopC) <> "" Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
    
    Else ' SH tiene efecto
          If UserList(Index).ConnID <> -1 Then
            Call HandleData(Index, rBuffer(LoopC))
          Else
            Exit Sub
          End If
    End If
        
Next LoopC

Exit Sub


ErrorHandler:
    Call LogError("Error en Socket read." & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)

#End If
End Sub



