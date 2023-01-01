VERSION 5.00
Begin VB.Form frmG_T_OPCIONES 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Configuración || Opciones (Opciones.ini)"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmG_T_OPCIONES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   4
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   24
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tPorcOro 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   150
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox tPorcExp 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   149
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   ".. de Oro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   153
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   ".. de Exp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   152
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Multiplicacion (NPC):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   151
         Top             =   120
         Width           =   2175
      End
      Begin VB.Shape Shape14 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CheckBox INISVR 
      BackColor       =   &H00404040&
      Caption         =   "Iniciar el Servidor al cerrar esta ventana..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   230
      Top             =   6000
      Visible         =   0   'False
      Width           =   9975
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   20
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tMantenimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   214
         Text            =   "5"
         Top             =   3360
         Width           =   375
      End
      Begin VB.CheckBox cParche 
         BackColor       =   &H00008000&
         Caption         =   "Utiliza Parche"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   212
         Top             =   3000
         Width           =   4215
      End
      Begin VB.CheckBox cAtributos011 
         BackColor       =   &H00008000&
         Caption         =   "Utilizar Atributos 0.11.x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   209
         Top             =   2640
         Width           =   4215
      End
      Begin VB.CheckBox cReservadoParaAdministradores 
         BackColor       =   &H00008000&
         Caption         =   "Reservado para Administradores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   203
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox cDecirConteo 
         BackColor       =   &H00008000&
         Caption         =   "Decir Conteo de Cerrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   201
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CheckBox cCerrarQuieto 
         BackColor       =   &H00008000&
         Caption         =   "Cerrar Quieto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   200
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox cAvisarGMs 
         BackColor       =   &H00008000&
         Caption         =   "Avisar si Inicia/Cierra un Dios/Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   198
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox tURLSoporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1680
         TabIndex        =   196
         Text            =   "http://ao.alkon.com.ar"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox cEstadisticasWebf 
         BackColor       =   &H00008000&
         Caption         =   "ESTADISTICAS ONLINE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label87 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "hs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   216
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label86 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Realizar Mantenimiento cada: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   215
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label Label83 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "URL Soporte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   197
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Servidor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   195
         Top             =   120
         Width           =   960
      End
      Begin VB.Shape Shape21 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   2
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   22
      Top             =   1560
      Width           =   10215
      Begin VB.CheckBox cPrivadoEnPantalla 
         BackColor       =   &H00008000&
         Caption         =   "Privado en Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   208
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox tModoAgarre 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   7800
         MaxLength       =   3
         TabIndex        =   205
         Text            =   "0"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1575
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   204
         Text            =   "frmG_T_OPCIONES.frx":1042
         ToolTipText     =   "Tabla de Configuraciones"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2775
         Left            =   2760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   184
         Text            =   "frmG_T_OPCIONES.frx":10F7
         ToolTipText     =   "Tabla de Configuraciones"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox tConfigClick 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   4680
         MaxLength       =   3
         TabIndex        =   183
         Text            =   "0"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox tConfigNPCClick 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   4680
         MaxLength       =   3
         TabIndex        =   182
         Text            =   "0"
         Top             =   3720
         Width           =   735
      End
      Begin VB.CheckBox cLluvia 
         BackColor       =   &H00008000&
         Caption         =   "Lluvia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   175
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox cVidaAlta 
         BackColor       =   &H00008000&
         Caption         =   "Vida Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox cBloquearPublicidades 
         BackColor       =   &H00008000&
         Caption         =   "Bloquear Publicidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   173
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox cNoHacerDiagnostico 
         BackColor       =   &H00008000&
         Caption         =   "No hacer diagnostico de errores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   172
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox NoVentanaDeInicioNW 
         BackColor       =   &H00008000&
         Caption         =   "No Mensaje NW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   171
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Modo Agarre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5880
         TabIndex        =   207
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Modos de Agarre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6000
         TabIndex        =   206
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Configurar los detalles:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   187
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label76 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "User Click:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   186
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "NPC Click:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   185
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Opciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   176
         Top             =   120
         Width           =   1065
      End
      Begin VB.Shape Shape16 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2385
      End
      Begin VB.Shape Shape18 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3855
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
      Begin VB.Shape Shape22 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2415
         Left            =   5760
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   6
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   26
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tSkillNavegacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   229
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tNivelNavegacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   226
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox cBajaStamina 
         BackColor       =   &H00008000&
         Caption         =   "Baja la Stamina Sin Ropa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   224
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CheckBox cHablanLosMuertos 
         BackColor       =   &H00008000&
         Caption         =   "Hablan los Muertos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   199
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox tMinBilletera 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   138
         Text            =   "0"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox cNoSeCaenLosItems 
         BackColor       =   &H00008000&
         Caption         =   "No Se Caen Los Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   137
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CheckBox cDesequiparAlMorir 
         BackColor       =   &H00008000&
         Caption         =   "Desequipar al Morir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   134
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox cEquiparAlRevivir 
         BackColor       =   &H00008000&
         Caption         =   "Equipar al Revivir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   133
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox cTirar100kAlMorir 
         BackColor       =   &H00008000&
         Caption         =   "Tirar Oro al Morir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   132
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox tExpKillUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   131
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label92 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Skills minimo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   228
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label91 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Nivel minimo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   227
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Navegacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   225
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Min. Billetera:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   139
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Usuarios:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   136
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   " Exp = Lvl por ... al matar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   135
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Shape Shape12 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3945
      End
      Begin VB.Shape Shape24 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   16
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   36
      Top             =   1560
      Width           =   10215
      Begin VB.CheckBox cPorNivel 
         BackColor       =   &H00008000&
         Caption         =   "x Nivel ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   223
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox tExpPorSkill 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   221
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox cSkillsRapidos 
         BackColor       =   &H00008000&
         Caption         =   "Ganar Skills naturales rapidamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   220
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd11x 
         BackColor       =   &H0000FF00&
         Caption         =   "Exp Lenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   1920
         Width           =   1755
      End
      Begin VB.CommandButton cmd99z 
         BackColor       =   &H0000FF00&
         Caption         =   "Exp 0.9.9z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   1920
         Width           =   1755
      End
      Begin VB.CheckBox cExperienciaRapida 
         BackColor       =   &H00008000&
         Caption         =   "Experiencia Rapida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   202
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox MenorNivel1 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   41
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox MenorNivel2 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox NivelMenorExp1 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   39
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox NivelMenorExp2 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   38
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox NivelDespuesExp 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   37
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label89 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Exp. por Skill:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   222
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Skills:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   219
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Despues usar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Experiencias:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Hasta nivel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Hasta nivel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "usar: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   43
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "usar: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.Shape Shape8 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2295
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3915
      End
      Begin VB.Shape Shape23 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   21
      Top             =   1560
      Width           =   10215
      Begin VB.CheckBox cAntiLukers 
         BackColor       =   &H00008000&
         Caption         =   "AntiLukers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   213
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox cAntiAOH 
         BackColor       =   &H00008000&
         Caption         =   "AntiAOh y otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   180
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox cAntiSpeedHack 
         BackColor       =   &H00008000&
         Caption         =   "AntiSpeedHack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   179
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox cAntiEntrenarZarpado 
         BackColor       =   &H00008000&
         Caption         =   "AntiEntrenar Zarpado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   178
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox cPermitirOcultarMensajes 
         BackColor       =   &H00008000&
         Caption         =   "Permitir ocultar mensajes..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   177
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Anti-Chits:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   181
         Top             =   120
         Width           =   1050
      End
      Begin VB.Shape Shape17 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2655
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   14
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   34
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tAltaHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   189
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox tMediaHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   64
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox tMinimaHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   63
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label78 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Alta hasta lvl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   190
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Media hasta lvl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Minima hasta lvl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Meditacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   65
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   12
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   32
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox Text3 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FF00&
         Height          =   1215
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   211
         Text            =   "frmG_T_OPCIONES.frx":12A5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox cNoSeCaenItemsEnTorneo 
         BackColor       =   &H00008000&
         Caption         =   "No se caen los items al morir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   210
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox tMapaDeTorneo 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   191
         Text            =   "0"
         Top             =   765
         Width           =   1575
      End
      Begin VB.TextBox tMaxMascotasTorneo 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   80
         Text            =   "0"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox tConfigTorneo 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   79
         Text            =   "0"
         Top             =   2460
         Width           =   1575
      End
      Begin VB.CheckBox cSePuedenUsarPots 
         BackColor       =   &H00008000&
         Caption         =   "Pot's Permitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CheckBox cNoKO 
         BackColor       =   &H00008000&
         Caption         =   "Anti-NoKO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Config. Torneo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   81
         Top             =   2470
         Width           =   1935
      End
      Begin VB.Label Label79 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mapa de Torneo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   192
         Top             =   765
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Torneos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   83
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Max. Mascotas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   82
         Top             =   420
         Width           =   1935
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3975
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3675
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   11
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   31
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tNivelFundarClan 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   84
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Nivel Fundar Clan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   85
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Clanes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   193
         Top             =   120
         Width           =   795
      End
      Begin VB.Shape Shape19 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   10
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   30
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tNivelLimiteNw 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   86
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Nivel Limite NW:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   87
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Newbies:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   194
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape20 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Guardar y Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   1755
   End
   Begin VB.CommandButton cmdDefecto 
      BackColor       =   &H0000FF00&
      Caption         =   "&Por defecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   1755
   End
   Begin VB.CommandButton Cerrar 
      BackColor       =   &H0000FF00&
      Caption         =   "&X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1755
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   15
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   35
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tInicioTTY 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   54
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox tModoCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   53
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox tInicioTTX 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   52
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tInicioCTX 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tInicioCTY 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   50
         Text            =   "0"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox tCS_GLD 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   49
         Text            =   "0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox tCS_Die 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   48
         Text            =   "0"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Criminal Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   62
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Modo Counter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   61
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mapa indicado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   60
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Criminal X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   59
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Ciudadano X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   58
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Ciudadano Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   57
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Valor minimo Ingreso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Valor de Muerte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   55
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3915
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   13
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   33
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tInicioAVX 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox tMapaAventura 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   70
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox tInicioAVY 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   69
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox tTiempoAV 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   68
         Text            =   "0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mapa indicado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   75
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Aventura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   74
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   73
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Tiempo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   72
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   9
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   29
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox CaosRecomp 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   93
         Text            =   "0"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox RealRecomp 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   92
         Text            =   "0"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox CiuCaos 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   91
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox CriReal 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   90
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox ExpRecompensa 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   89
         Text            =   "0"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox ExpEnlistar 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   88
         Text            =   "0"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Facciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   100
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label17 
         BackColor       =   &H0000C000&
         Caption         =   "Exp. de Enlistarse:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   99
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label15 
         BackColor       =   &H0000C000&
         Caption         =   "Exp. de Recompensa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   98
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000C000&
         Caption         =   "Cri. para Armada Real:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000C000&
         Caption         =   "Ciu. para Ejercito Caos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   96
         Top             =   960
         Width           =   2580
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000C000&
         Caption         =   "Recompensa Real:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   95
         Top             =   1800
         Width           =   2610
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000C000&
         Caption         =   "Recompensa Caos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   94
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3945
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   8
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   28
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox SubNivSx 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   104
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox SubNivSm 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   103
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox NuevoAx 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   102
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox NuevoAm 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   101
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Atributos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   113
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Al subir de Nivel dar...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   112
         Top             =   600
         Width           =   2370
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "de:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   111
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Skill's a "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   110
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Skill's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   109
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Al crear un nuevo Personaje tirar..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   108
         Top             =   1560
         Width           =   3585
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "de:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   107
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Dados a "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1350
         TabIndex        =   106
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2895
         TabIndex        =   105
         Top             =   1920
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   3585
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   3585
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   2295
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   3945
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   7
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   27
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tMaxOro 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2400
         TabIndex        =   121
         Text            =   "0"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox tMaxExp 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2400
         TabIndex        =   120
         Text            =   "0"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox MaxEnergia 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   119
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox MaxMana 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   118
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox MaxVida 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   117
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox MaxNivel 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   116
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox tMaxHit 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   115
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox tMaxDef 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   114
         Text            =   "0"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Maximos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   130
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label51 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Exp.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   129
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label50 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Oro portable:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   128
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label49 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Nivel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   127
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label48 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Vida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   126
         Top             =   840
         Width           =   3195
      End
      Begin VB.Label Label47 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Mana:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   125
         Top             =   1200
         Width           =   3045
      End
      Begin VB.Label Label46 
         BackColor       =   &H0000C000&
         Caption         =   "Max. Energia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   124
         Top             =   1560
         Width           =   3285
      End
      Begin VB.Label Label45 
         BackColor       =   &H0000C000&
         Caption         =   "Max. DEF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   123
         Top             =   2280
         Width           =   3285
      End
      Begin VB.Label Label44 
         BackColor       =   &H0000C000&
         Caption         =   "Max. HIT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   122
         Top             =   1920
         Width           =   3045
      End
      Begin VB.Shape Shape11 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4305
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   5
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   25
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tResMinHP 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   143
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox tResMaxHP 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   142
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tResMinMP 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   141
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox tResMaxMP 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3360
         MaxLength       =   9
         TabIndex        =   140
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Resucitar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   148
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Min. HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   147
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Max. HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   146
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Min. MP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   145
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Max. MP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   144
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Shape Shape13 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4305
      End
   End
   Begin VB.PictureBox CONF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   3
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   10185
      TabIndex        =   23
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox tLoteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   161
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tOtracara 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   160
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox tAventura 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   159
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox tMoverUlla 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   158
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tMoverBander 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   157
         Text            =   "0"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox tMoverNix 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   156
         Text            =   "0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox tMoverLindos 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   155
         Text            =   "0"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox tMoverVeril 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   154
         Text            =   "0"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Boleto de Loteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   170
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Reconstructor Facial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   169
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Precios:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   168
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label69 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Boleto de Aventura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   167
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mover a Ullathorpe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   166
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mover a Banderbill:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   165
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label66 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mover a Nix:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   164
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mover a Lindos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   163
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Mover a Veril:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   162
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Shape Shape15 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3915
      End
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "EXPERIENCIAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   6720
      TabIndex        =   16
      Top             =   960
      Width           =   1650
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "COUNTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   5400
      TabIndex        =   15
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "MEDITACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   3840
      TabIndex        =   14
      Top             =   960
      Width           =   1410
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "AVENTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   2400
      TabIndex        =   13
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "TORNEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   1320
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "CLANES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   900
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "NEWBIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   7080
      TabIndex        =   10
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "FACCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   5640
      TabIndex        =   9
      Top             =   600
      Width           =   1290
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "ATRIBUTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "MAXIMOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   3000
      TabIndex        =   7
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "USUARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "RESUCITAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "MULTIPLICACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1860
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "PRECIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "OPCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "ANTI-CHITS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label OPCION 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "SERVIDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8400
   End
End
Attribute VB_Name = "frmG_T_OPCIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub VaciarTodo()
Dim i As Integer
For i = 0 To CONF.Count - 1
    CONF(i).Visible = False
    OPCION(i).BackColor = &HC000&
Next

End Sub


Private Sub Cerrar_Click()
Unload Me
End Sub



Private Sub cmd11x_Click()
MenorNivel1 = 8
NivelMenorExp1 = 1.3
MenorNivel2 = 24
NivelMenorExp2 = 1.2
NivelDespuesExp = 1.1
cExperienciaRapida.Value = 0

End Sub

Private Sub cmd99z_Click()
MenorNivel1 = 11
NivelMenorExp1 = 1.5
MenorNivel2 = 25
NivelMenorExp2 = 1.3
NivelDespuesExp = 1.2
cExperienciaRapida.Value = 0

End Sub

Private Sub Form_Load()
Call VaciarTodo
Call cmdDefecto_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Me.Tag = "UNICO" Then
    If INISVR.Visible = True And INISVR.Value = 1 Then Call Shell(App.Path & "\" & App.EXEName & ".exe -ejecutarigual", vbNormalFocus)
    End
End If
End Sub

Private Sub MenorNivel1_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub MenorNivel2_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub NivelDespuesExp_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
End Sub

Private Sub NivelMenorExp1_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
End Sub

Private Sub NivelMenorExp2_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
End Sub

Private Sub OPCION_Click(Index As Integer)
Call VaciarTodo
OPCION(Index).BackColor = vbGreen
CONF(Index).Visible = True
End Sub


Private Sub cmdDefecto_Click()
On Error Resume Next

' 0.12b3
tExpPorSkill = ExpPorSkill
If PorNivel = True Then
    cPorNivel.Value = 1
Else
    cPorNivel.Value = 0
End If
If BajaStamina = True Then
    cBajaStamina.Value = 1
Else
    cBajaStamina.Value = 0
End If
tNivelNavegacion.Text = NivelNavegacion
tSkillNavegacion.Text = SkillNavegacion


If AntiAOH = True Then
    cAntiAOH.Value = 1
Else
    cAntiAOH.Value = 0
End If

tMantenimiento.Text = HsMantenimientoReal

If SkillsRapidos = True Then
    cSkillsRapidos.Value = 1
Else
    cSkillsRapidos.Value = 0
End If

cParche.Value = Parche

If AntiLukers = True Then
    cAntiLukers.Value = 1
Else
    cAntiLukers.Value = 0
End If

If AntiSpeedHack = True Then
    cAntiSpeedHack.Value = 1
Else
    cAntiSpeedHack.Value = 0
End If
If AntiEntrenar = True Then
    cAntiEntrenarZarpado.Value = 1
Else
    cAntiEntrenarZarpado.Value = 0
End If
If PermitirOcultarMensajes = True Then
    cPermitirOcultarMensajes.Value = 1
Else
    cPermitirOcultarMensajes.Value = 0
End If

If NoSeCaenItemsEnTorneo = True Then
    cNoSeCaenItemsEnTorneo.Value = 1
Else
    cNoSeCaenItemsEnTorneo.Value = 0
End If

If Atributos011 = True Then
    cAtributos011.Value = 1
Else
    cAtributos011.Value = 0
End If

If ReservadoParaAdministradores = True Then
    cReservadoParaAdministradores.Value = 1
Else
    cReservadoParaAdministradores.Value = 0
End If

If PrivadoEnPantalla = True Then
    cPrivadoEnPantalla.Value = 1
Else
    cPrivadoEnPantalla.Value = 0
End If
    
If DecirConteo = True Then
    cDecirConteo.Value = 1
Else
    cDecirConteo.Value = 0
End If

If CerrarQuieto = True Then
    cCerrarQuieto.Value = 1
Else
    cCerrarQuieto.Value = 0
End If

If Muertos_Hablan = True Then
    cHablanLosMuertos.Value = 1
Else
    cHablanLosMuertos.Value = 0
End If

If NoMensajeANW = True Then
    NoVentanaDeInicioNW.Value = 1
Else
    NoVentanaDeInicioNW.Value = 0
End If

If EscrachGM = True Then
    cAvisarGMs.Value = 1
Else
    cAvisarGMs.Value = 0
End If

If ExperienciaRapida = True Then
    cExperienciaRapida.Value = 1
Else
    cExperienciaRapida.Value = 0
End If

If LluviaON = True Then
    cLluvia.Value = 1
Else
    cLluvia.Value = 0
End If
If VidaAlta = True Then
    cVidaAlta.Value = 1
Else
    cVidaAlta.Value = 0
End If
If Publicidad = True Then
    cBloquearPublicidades.Value = 1
Else
    cBloquearPublicidades.Value = 0
End If
If NoHacerDiagnosticoDeErrores = True Then
    cNoHacerDiagnostico.Value = 1
Else
    cNoHacerDiagnostico.Value = 0
End If

If PotsEnTorneo = True Then
    cSePuedenUsarPots.Value = 1
Else
    cSePuedenUsarPots.Value = 0
End If

If DesequiparAlMorir = True Then
    cDesequiparAlMorir.Value = 1
Else
    cDesequiparAlMorir.Value = 0
End If
If EquiparAlRevivir = True Then
    cEquiparAlRevivir.Value = 1
Else
    cEquiparAlRevivir.Value = 0
End If
If Tirar100kAlMorir = True Then
    cTirar100kAlMorir.Value = 1
Else
    cTirar100kAlMorir.Value = 0
End If

If NoKO = True Then
    cNoKO.Value = 1
Else
    cNoKO.Value = 0
End If

tMapaDeTorneo.Text = MapaDeTorneo

tResMinHP.Text = ResMinHP
tResMaxHP.Text = ResMaxHP
tResMinMP.Text = ResMinMP
tResMaxMP.Text = ResMaxMP

tURLSoporte.Text = URL_Soporte

tModoAgarre.Text = ModoAgarre

tExpKillUser.Text = ExpKillUser

tNivelLimiteNw.Text = LimiteNewbie
tNivelFundarClan.Text = NivelMinimoParaFundar
tMaxMascotasTorneo.Text = MaxMascotasTorneo
tConfigTorneo = ConfigTorneo

ExpEnlistar = ExpAlUnirse
ExpRecompensa = ExpX100
CriReal = ParaArmada
CiuCaos = ParaCaos
RealRecomp = RecompensaXArmada
CaosRecomp = RecompensaXCaos
tMaxExp = MaxExp
tMaxOro = MaxOro
MaxNivel = STAT_MAXELV
MaxVida = STAT_MAXHP
MaxMana = STAT_MAXMAN
tMaxHit = STAT_MAXHIT
tMaxDef = STAT_MAXDEF
MaxEnergia = STAT_MAXSTA
SubNivSm = MINSKILL_G
SubNivSx = MAXSKILL_G
NuevoAm = MINATTRB
NuevoAx = MAXATTRB
DoEvents


MenorNivel1 = Exp_MenorQ1
MenorNivel2 = Exp_MenorQ2
NivelMenorExp1 = Exp_Menor1
NivelMenorExp2 = Exp_Menor2
NivelDespuesExp = Exp_Despues
tOtracara = ReconstructorFacial
tLoteria = BoletoDeLoteria
tAventura = BoletoAventura
tMoverUlla = MoverUlla
tMoverBander = MoverBander
tMoverNix = MoverNix
tMoverLindos = MoverLindos
tMoverVeril = MoverVeril
tMapaAventura = MapaAventura
tInicioAVX = InicioAVX
tInicioAVY = InicioAVY
tTiempoAV = TiempoAV
tPorcOro = PorcORO
tPorcExp = PorcEXP

tMinimaHasta = MeditarChicoHasta
tMediaHasta = MeditarMedioHasta
tAltaHasta = MeditarAltaHasta

tModoCounter = MapaCounter
tInicioTTY = InicioTTY
tInicioTTX = InicioTTX
tInicioCTY = InicioCTY
tInicioCTX = InicioCTX
tCS_GLD = CS_GLD
tCS_Die = CS_Die

tConfigClick = ConfigClick
tConfigNPCClick = ConfigNPCClick
tMinBilletera = MinBilletera
If EstadisticasWebF = True Then
    cEstadisticasWebf.Value = 1
Else
    cEstadisticasWebf.Value = 0
End If
If NoSeCaenItems = True Then
    cNoSeCaenLosItems.Value = 1
Else
    cNoSeCaenLosItems.Value = 0
End If
End Sub

Private Sub cmdGuardar_Click()
On Error Resume Next

Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiAOH", cAntiAOH.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiSpeedHack", cAntiSpeedHack.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiEntrenarZarpado", cAntiEntrenarZarpado.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "PermitirOcultarMensajes", cPermitirOcultarMensajes.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiLukers", cAntiLukers.Value)

Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "BloqPublicidad", cPermitirOcultarMensajes.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "Lluvia", cLluvia.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "VidaAlta", cVidaAlta.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "NOHacerDiagnostico", cNoHacerDiagnostico.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "NoVentanaDeInicioNW", NoVentanaDeInicioNW.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "PrivadoEnPantalla", cPrivadoEnPantalla.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ModoAgarre", val(tModoAgarre.Text))
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ConfigClick", val(tConfigClick.Text))
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ConfigNPCClick", val(tConfigNPCClick.Text))

Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MinHP", val(tResMinHP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxHP", val(tResMaxHP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MinMP", val(tResMinMP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxMP", val(tResMaxMP))

Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "ExpKillUser", val(tExpKillUser))
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "DesequiparAlMorir", cDesequiparAlMorir.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "EquiparAlRevivir", cEquiparAlRevivir.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "Tirar100kAlMorir", cTirar100kAlMorir.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "MuertosHablan", cHablanLosMuertos.Value)
' 0.12b3
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "BajaStamina", cBajaStamina.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "SkillNavegacion", val(tSkillNavegacion))
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "NivelNavegacion", val(tNivelNavegacion))


Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "ParaCaos", val(CiuCaos))
Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "ParaArmada", val(CriReal))
Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaArmada", val(RealRecomp))
Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaCaos", val(CaosRecomp))
Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "EnlistarExp", val(ExpEnlistar))
Call WriteVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaExp", val(ExpRecompensa))

Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxLVL", val(MaxNivel))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxHP", val(MaxVida))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxMAN", val(MaxMana))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxST", val(MaxEnergia))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxMAN", val(MaxMana))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxHIT", val(tMaxHit))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxDEF", val(tMaxDef))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxEXP", val(tMaxExp))
Call WriteVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxORO", val(tMaxOro))

Call WriteVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MinSKILL", val(SubNivSm))
Call WriteVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MaxSKILL", val(SubNivSx))
Call WriteVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MinAtrib", val(NuevoAm))
Call WriteVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MaxAtrib", val(NuevoAx))

Call WriteVar(IniPath & "Opciones.ini", "NEWBIES", "NivelLimiteNw", val(tNivelLimiteNw))

Call WriteVar(IniPath & "Opciones.ini", "CLANES", "NivelMinimoParaFundar", val(tNivelFundarClan))

Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "ConfigTorneo", val(tConfigTorneo))
Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "MaxMascotasTorneo", val(tMaxMascotasTorneo))
Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "ValenPots", cSePuedenUsarPots.Value)
Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "NoKO", cNoKO.Value)
Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "MapaDeTorneo", val(tMapaDeTorneo))
Call WriteVar(IniPath & "Opciones.ini", "TORNEO", "NoSeCaenItemsEnTorneo", cNoSeCaenItemsEnTorneo.Value)

NivelMenorExp1.Text = CStr(Replace(NivelMenorExp1.Text, ",", "."))
NivelMenorExp2.Text = CStr(Replace(NivelMenorExp2.Text, ",", "."))
NivelDespuesExp.Text = CStr(Replace(NivelDespuesExp.Text, ",", "."))

Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor1", MenorNivel1.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor2", MenorNivel2.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp1", NivelMenorExp1.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp2", NivelMenorExp2.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelDespuesExp", NivelDespuesExp.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "ExperienciaRapida", cExperienciaRapida.Value)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "SkillsRapidos", cSkillsRapidos.Value)
' 0.12b3
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "ExpPorSkill", tExpPorSkill.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "PorNivel", cPorNivel.Value)

Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "EstadisticasWeb", cEstadisticasWebf.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "AvisarGMs", cAvisarGMs.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "URLSoporte", tURLSoporte.Text)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "CerrarQuieto", cCerrarQuieto.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "DecirConteoDeCerrado", cDecirConteo.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "Atributos011", cAtributos011.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "ReservadoParaAdministradores", cReservadoParaAdministradores.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "Parche", cParche.Value)
Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "Mantenimiento", val(tMantenimiento.Text))

Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "MinBilletera", val(tMinBilletera))
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "NoSeCaenLosItems", cNoSeCaenLosItems.Value)

Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "ReconstructorFacial", val(tOtracara))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "BoletoDeLoteria", val(tLoteria))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "BoletoAventura", val(tAventura))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "MoverUlla", val(tMoverUlla))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "MoverBander", val(tMoverBander))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "MoverNix", val(tMoverNix))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "MoverLindos", val(tMoverLindos))
Call WriteVar(IniPath & "Opciones.ini", "PRECIOS", "MoverVeril", val(tMoverVeril))

Call WriteVar(IniPath & "Opciones.ini", "MEDITACION", "MinimaHasta", val(tMinimaHasta.Text))
Call WriteVar(IniPath & "Opciones.ini", "MEDITACION", "MediaHasta", val(tMediaHasta.Text))
Call WriteVar(IniPath & "Opciones.ini", "MEDITACION", "AltaHasta", val(tAltaHasta.Text))

tPorcOro = Replace(tPorcOro, ",", ".")
tPorcExp = Replace(tPorcExp, ",", ".")

Call WriteVar(IniPath & "Opciones.ini", "MULTIPLICACION", "Oro", val(tPorcOro))
Call WriteVar(IniPath & "Opciones.ini", "MULTIPLICACION", "Exp", val(tPorcExp))

Call WriteVar(IniPath & "Opciones.ini", "AVENTURA", "MapaAventura", val(tMapaAventura))
Call WriteVar(IniPath & "Opciones.ini", "AVENTURA", "InicioX", val(tInicioAVX))
Call WriteVar(IniPath & "Opciones.ini", "AVENTURA", "InicioY", val(tInicioAVY))
Call WriteVar(IniPath & "Opciones.ini", "AVENTURA", "TiempoAventura", val(tTiempoAV))

Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "MapaCounter", val(tModoCounter))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "IniCriX", val(tInicioTTX))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "IniCriY", val(tInicioTTY))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "IniCiuX", val(tInicioCTX))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "IniCiuY", val(tInicioCTY))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "IngresoMinimo", val(tCS_GLD))
Call WriteVar(IniPath & "Opciones.ini", "COUNTER", "ValorMuerte", val(tCS_Die))
DoEvents
LoadOpcsINI
Unload Me
End Sub

Private Sub Text4_Change()

End Sub

Private Sub tMantenimiento_LostFocus()

If IsNumeric(tMantenimiento.Text) = False Then
    tMantenimiento.Text = HsMantenimientoReal
Else
    If tMantenimiento.Text < 5 Then tMantenimiento.Text = 5
    If tMantenimiento.Text > 24 Then tMantenimiento.Text = 24
End If

If tMantenimiento.Text <> HsMantenimientoReal Then
    MsgBox "Los cambios de tiempo del Mantenimiento se realizaran en el proximo Mantenimiento."
End If

End Sub
