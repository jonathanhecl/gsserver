VERSION 5.00
Begin VB.Form frmG_C_Opciones_2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Opciones 2da Parte"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "frmG_C_Opciones_2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   10485
   Begin VB.TextBox NivelDespuesExp 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   74
      Text            =   "0"
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox NivelMenorExp2 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   73
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox NivelMenorExp1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   72
      Text            =   "0"
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox MenorNivel2 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   67
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox MenorNivel1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   66
      Text            =   "0"
      Top             =   6960
      Width           =   735
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
      Height          =   375
      Left            =   4200
      TabIndex        =   65
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox tConfigNPCClick 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9360
      MaxLength       =   3
      TabIndex        =   63
      Text            =   "0"
      Top             =   4560
      Width           =   735
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
      Left            =   4320
      TabIndex        =   62
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox tMinBilletera 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5760
      MaxLength       =   9
      TabIndex        =   59
      Text            =   "0"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox tConfigClick 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9360
      MaxLength       =   3
      TabIndex        =   57
      Text            =   "0"
      Top             =   4200
      Width           =   735
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
      Height          =   3495
      Left            =   7440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   56
      Text            =   "frmG_C_Opciones_2.frx":000C
      ToolTipText     =   "Tabla de Configuraciones"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox tCS_Die 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   54
      Text            =   "0"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox tCS_GLD 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   51
      Text            =   "0"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox tInicioCTY 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   49
      Text            =   "0"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox tInicioCTX 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   47
      Text            =   "0"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox tInicioTTX 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   42
      Text            =   "0"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox tModoCounter 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   41
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox tInicioTTY 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   40
      Text            =   "0"
      Top             =   4680
      Width           =   1215
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox tMoverVeril 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   39
      Text            =   "0"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox tMoverLindos 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   37
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox tMoverNix 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   35
      Text            =   "0"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox tMoverBander 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   33
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tMoverUlla 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   31
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox tAventura 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   30
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox tOtracara 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   24
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox tLoteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   23
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox tTiempoAV 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   22
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox tInicioAVY 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   21
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox tMapaAventura 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox tInicioAVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox tPorcExp 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   10
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox tPorcOro 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   9
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox tMinimaHasta 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   5
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox tMediaHasta 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   4
      Text            =   "0"
      Top             =   2160
      Width           =   735
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   2775
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "&Opciones 1ra Parte"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   7800
      Picture         =   "frmG_C_Opciones_2.frx":01BA
      Top             =   5400
      Width           =   2175
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
      TabIndex        =   76
      Top             =   7320
      Width           =   855
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
      TabIndex        =   75
      Top             =   6960
      Width           =   855
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
      TabIndex        =   71
      Top             =   7320
      Width           =   1455
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
      TabIndex        =   70
      Top             =   6960
      Width           =   1455
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
      TabIndex        =   69
      Top             =   6600
      Width           =   1425
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
      TabIndex        =   68
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label31 
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
      Left            =   7440
      TabIndex        =   64
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label30 
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
      Left            =   4320
      TabIndex        =   61
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Usuarios (Otros)"
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
      TabIndex        =   60
      Top             =   4560
      Width           =   1710
   End
   Begin VB.Label Label28 
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
      Left            =   7440
      TabIndex        =   58
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label27 
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
      Left            =   7560
      TabIndex        =   55
      Top             =   240
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
      TabIndex        =   53
      Top             =   6120
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
      TabIndex        =   52
      Top             =   5760
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
      TabIndex        =   50
      Top             =   5400
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
      TabIndex        =   48
      Top             =   5040
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
      TabIndex        =   46
      Top             =   4320
      Width           =   2415
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
      TabIndex        =   45
      Top             =   3960
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
      TabIndex        =   44
      Top             =   3600
      Width           =   1515
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
      TabIndex        =   43
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label18 
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
      TabIndex        =   38
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label17 
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
      TabIndex        =   36
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label15 
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
      TabIndex        =   34
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label14 
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
      TabIndex        =   32
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label13 
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
      TabIndex        =   29
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label12 
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
      TabIndex        =   28
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label11 
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
      TabIndex        =   27
      Top             =   240
      Width           =   870
   End
   Begin VB.Label Label10 
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
      TabIndex        =   26
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label9 
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
      TabIndex        =   25
      Top             =   960
      Width           =   2415
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
      Left            =   4320
      TabIndex        =   20
      Top             =   4080
      Width           =   1935
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
      Left            =   4320
      TabIndex        =   19
      Top             =   3720
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
      Left            =   4440
      TabIndex        =   18
      Top             =   2640
      Width           =   990
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
      Left            =   4320
      TabIndex        =   17
      Top             =   3000
      Width           =   1935
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
      Left            =   4320
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Porcentajes (NPC):"
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
      TabIndex        =   13
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label2 
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
      Left            =   4320
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Left            =   4320
      TabIndex        =   11
      Top             =   960
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
      Left            =   4440
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   4200
      Top             =   6960
      Width           =   2985
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   4200
      Top             =   6240
      Width           =   2985
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   4200
      Top             =   1560
      Width           =   2955
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   4200
      Top             =   360
      Width           =   2955
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   4200
      Top             =   2760
      Width           =   2955
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   120
      Top             =   360
      Width           =   3915
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   120
      Top             =   3720
      Width           =   3915
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   7320
      Top             =   360
      Width           =   2955
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   4200
      Top             =   4680
      Width           =   2955
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Top             =   6720
      Width           =   3915
   End
End
Attribute VB_Name = "frmG_C_Opciones_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
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

Private Sub Command2_Click()
On Error Resume Next
'ReconstructorFacial = 200000
'BoletoDeLoteria = 20000
'BoletoAventura = 20000
'MoverUlla = 9999
'MoverBander = 9999
'MoverLindos = 9999
'MoverNix = 9999
'MoverVeril = 200000

'MenorNivel1 = Exp_MenorQ1
'MenorNivel2 = Exp_MenorQ2
'NivelMenorExp1 = Exp_Menor1
'NivelMenorExp2 = Exp_Menor2
'NivelDespuesExp = Exp_Despues

Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor1", MenorNivel1.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor2", MenorNivel2.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp1", NivelMenorExp1.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp2", NivelMenorExp2.Text)
Call WriteVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelDespuesExp", NivelDespuesExp.Text)

Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "EstadisticasWeb", cEstadisticasWebf.Value)

Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ConfigClick", val(tConfigClick))
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ConfigNPCClick", val(tConfigNPCClick))

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

Call WriteVar(IniPath & "Opciones.ini", "MEDITACION", "MinimaHasta", val(tMinimaHasta))
Call WriteVar(IniPath & "Opciones.ini", "MEDITACION", "MediaHasta", val(tMediaHasta))

Call WriteVar(IniPath & "Opciones.ini", "PORCENTAJES", "Oro", val(tPorcOro))
Call WriteVar(IniPath & "Opciones.ini", "PORCENTAJES", "Exp", val(tPorcExp))

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

LoadOpcsINI
Unload Me
End Sub

Private Sub Command3_Click()
frmG_C_Opciones.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = frmGeneral.Icon
Me.Left = 0
Me.Top = 0
Call Command1_Click
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

