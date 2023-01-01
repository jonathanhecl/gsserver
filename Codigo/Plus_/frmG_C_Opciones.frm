VERSION 5.00
Begin VB.Form frmG_C_Opciones 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Opciones 1ra Parte"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "frmG_C_Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11190
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
      Height          =   375
      Left            =   240
      TabIndex        =   84
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "&Opciones 2da Parte"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CheckBox cNoKO 
      BackColor       =   &H00008000&
      Caption         =   "No KO"
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
      Left            =   9600
      TabIndex        =   82
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox tResMaxMP 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   80
      Text            =   "0"
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox tResMinMP 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   78
      Text            =   "0"
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox tResMaxHP 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   76
      Text            =   "0"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox tResMinHP 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   74
      Text            =   "0"
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox tExpKillUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5520
      MaxLength       =   9
      TabIndex        =   71
      Text            =   "0"
      Top             =   7200
      Width           =   855
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
      Left            =   2880
      TabIndex        =   70
      Top             =   6840
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
      Left            =   2880
      TabIndex        =   69
      Top             =   6480
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
      Left            =   2880
      TabIndex        =   68
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CheckBox cSePuedenUsarPots 
      BackColor       =   &H00008000&
      Caption         =   "Si Pot's"
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
      Left            =   9600
      TabIndex        =   66
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox tConfigTorneo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      MaxLength       =   9
      TabIndex        =   65
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox tMaxMascotasTorneo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      MaxLength       =   9
      TabIndex        =   62
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox tNivelFundarClan 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      MaxLength       =   9
      TabIndex        =   59
      Text            =   "0"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox tNivelLimiteNw 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      MaxLength       =   9
      TabIndex        =   57
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox tMaxDef 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   54
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox tMaxHit 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   53
      Text            =   "0"
      Top             =   1920
      Width           =   855
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
      TabIndex        =   52
      Top             =   4440
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
      TabIndex        =   51
      Top             =   3840
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
      TabIndex        =   50
      Top             =   3480
      Width           =   2055
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
      TabIndex        =   49
      Top             =   3120
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
      TabIndex        =   47
      Top             =   1800
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
      TabIndex        =   46
      Top             =   1200
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
      TabIndex        =   45
      Top             =   840
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
      TabIndex        =   44
      Top             =   480
      Width           =   2055
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
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox NuevoAm 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   34
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox NuevoAx 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   33
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox MaxNivel 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   27
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox MaxVida 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   26
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox MaxMana 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   25
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox MaxEnergia 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9960
      MaxLength       =   9
      TabIndex        =   24
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox tMaxExp 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9000
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox tMaxOro 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9000
      TabIndex        =   22
      Text            =   "0"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox ExpEnlistar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   16
      TabIndex        =   14
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox ExpRecompensa 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   16
      TabIndex        =   13
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox CriReal 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox CiuCaos 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   11
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox RealRecomp 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   10
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox CaosRecomp 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   9
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox SubNivSm 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox SubNivSx 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   5880
      Width           =   2385
   End
   Begin VB.Label Label43 
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
      Left            =   6840
      TabIndex        =   81
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label42 
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
      Left            =   6840
      TabIndex        =   79
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label41 
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
      Left            =   6840
      TabIndex        =   77
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label40 
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
      Left            =   6840
      TabIndex        =   75
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label39 
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
      Left            =   6960
      TabIndex        =   73
      Top             =   5760
      Width           =   1065
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   6720
      Top             =   5880
      Width           =   4305
   End
   Begin VB.Label Label38 
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
      Left            =   2880
      TabIndex        =   72
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label37 
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
      Left            =   2880
      TabIndex        =   67
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Label Label36 
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
      Left            =   6840
      TabIndex        =   64
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label35 
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
      Left            =   6840
      TabIndex        =   63
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label34 
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
      Left            =   6960
      TabIndex        =   61
      Top             =   4620
      Width           =   945
   End
   Begin VB.Label Label33 
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
      Left            =   6840
      TabIndex        =   60
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label32 
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
      Left            =   6840
      TabIndex        =   58
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label31 
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
      Left            =   6960
      TabIndex        =   56
      Top             =   1920
      Width           =   3045
   End
   Begin VB.Label Label30 
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
      Left            =   6960
      TabIndex        =   55
      Top             =   2280
      Width           =   3285
   End
   Begin VB.Label Label17 
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
      TabIndex        =   48
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   120
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Label Label2 
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
      TabIndex        =   43
      Top             =   120
      Width           =   1050
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   120
      Top             =   240
      Width           =   2385
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Top             =   6600
      Width           =   2385
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Otros:"
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
      Left            =   6960
      TabIndex        =   39
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label15 
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
      Left            =   5415
      TabIndex        =   38
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label14 
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
      Left            =   3870
      TabIndex        =   37
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label13 
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
      Left            =   2880
      TabIndex        =   36
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label Label12 
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
      Left            =   2760
      TabIndex        =   35
      Top             =   4680
      Width           =   3585
   End
   Begin VB.Label Label11 
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
      Left            =   5400
      TabIndex        =   32
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label Label10 
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
      Left            =   3840
      TabIndex        =   31
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Label Label9 
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
      Left            =   2880
      TabIndex        =   30
      Top             =   4080
      Width           =   465
   End
   Begin VB.Label Label8 
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
      Left            =   2880
      TabIndex        =   29
      Top             =   3720
      Width           =   2370
   End
   Begin VB.Label Label7 
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
      Left            =   2880
      TabIndex        =   28
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Label28 
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
      Left            =   6960
      TabIndex        =   21
      Top             =   1560
      Width           =   3285
   End
   Begin VB.Label Label27 
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
      Left            =   6960
      TabIndex        =   20
      Top             =   1200
      Width           =   3045
   End
   Begin VB.Label Label6 
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
      Left            =   6960
      TabIndex        =   19
      Top             =   840
      Width           =   3195
   End
   Begin VB.Label Label5 
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
      Left            =   6960
      TabIndex        =   18
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label4 
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
      Left            =   6960
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   6960
      TabIndex        =   16
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Left            =   6960
      TabIndex        =   15
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label26 
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
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label25 
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
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Width           =   2610
   End
   Begin VB.Label Label24 
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
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   2580
   End
   Begin VB.Label Label23 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   2505
   End
   Begin VB.Label Label22 
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
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label21 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label20 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   2640
      Top             =   240
      Width           =   3945
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   6720
      Top             =   240
      Width           =   4305
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2760
      Top             =   3840
      Width           =   3585
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2760
      Top             =   4800
      Width           =   3585
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   2640
      Top             =   3380
      Width           =   3945
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   6720
      Top             =   3600
      Width           =   4275
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   6720
      Top             =   4690
      Width           =   4275
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   2640
      Top             =   5880
      Width           =   3945
   End
End
Attribute VB_Name = "frmG_C_Opciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
DoEvents
If AntiAOH = True Then
    cAntiAOH.Value = 1
Else
    cAntiAOH.Value = 0
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

If NoMensajeANW = True Then
    NoVentanaDeInicioNW.Value = 1
Else
    NoVentanaDeInicioNW.Value = 0
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

tResMinHP.Text = ResMinHP
tResMaxHP.Text = ResMaxHP
tResMinMP.Text = ResMinMP
tResMaxMP.Text = ResMaxMP

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
End Sub

Private Sub Command2_Click()
On Error Resume Next
'If Not Numeric(Text1.Text) Or Not Numeric(Text2.Text) Or Not Numeric(Text3.Text) Or Not Numeric(Text4.Text) Or Not Numeric(Text5.Text) Or Not Numeric(Text6.Text) Or Not Numeric(Text7.Text) Or Not Numeric(Text8.Text) Or Not Numeric(Text9.Text) Or Not Numeric(Text10.Text) Or Not Numeric(Text11.Text) Or Not Numeric(Text12.Text) Or Not Numeric(Text13.Text) Or Not Numeric(Text14.Text) Or Not Numeric(Text15.Text) Or Not Numeric(Text16.Text) Or Not Numeric(Text17.Text) Then
'    MsgBox "Alguno de los campos ingresados, no contiene un valor numerico."
'    Exit Sub
'End If


Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiAOH", cAntiAOH.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiSpeedHack", cAntiSpeedHack.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiEntrenarZarpado", cAntiEntrenarZarpado.Value)
Call WriteVar(IniPath & "Opciones.ini", "ANTI-CHITS", "PermitirOcultarMensajes", cPermitirOcultarMensajes.Value)

Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "BloqPublicidad", cPermitirOcultarMensajes.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "Lluvia", cLluvia.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "VidaAlta", cVidaAlta.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "NOHacerDiagnostico", cNoHacerDiagnostico.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "NoVentanaDeInicioNW", NoVentanaDeInicioNW.Value)

Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MinHP", val(tResMinHP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxHP", val(tResMaxHP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MinMP", val(tResMinMP))
Call WriteVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxMP", val(tResMaxMP))

Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "ExpKillUser", val(tExpKillUser))
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "DesequiparAlMorir", cDesequiparAlMorir.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "EquiparAlRevivir", cEquiparAlRevivir.Value)
Call WriteVar(IniPath & "Opciones.ini", "USUARIOS", "Tirar100kAlMorir", cTirar100kAlMorir.Value)

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


LoadOpcsINI
Unload Me
End Sub

Private Sub Command3_Click()
frmG_C_Opciones_2.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = frmGeneral.Icon
Me.Left = 0
Me.Top = 0
Call Command1_Click
End Sub

