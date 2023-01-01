VERSION 5.00
Begin VB.Form frmG_C_Intervalos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Configuración || Intervalos (Server.ini)"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmG_C_Intervalos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   7785
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Aplictar y &Guardar"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtWS 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   240
      TabIndex        =   58
      Text            =   "0"
      Top             =   5280
      Width           =   900
   End
   Begin VB.CommandButton cmdDefecto 
      BackColor       =   &H0000FF00&
      Caption         =   "&Por Defecto"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4920
      Width           =   1695
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
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtNPCPuedeAtacar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   5640
      TabIndex        =   51
      Text            =   "0"
      Top             =   3345
      Width           =   1635
   End
   Begin VB.TextBox txtAI 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   5640
      TabIndex        =   50
      Text            =   "0"
      Top             =   4005
      Width           =   1635
   End
   Begin VB.TextBox txtIntervaloWavFx 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   3000
      TabIndex        =   43
      Text            =   "0"
      Top             =   3345
      Width           =   795
   End
   Begin VB.TextBox txtIntervaloFrio 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   3000
      TabIndex        =   42
      Text            =   "0"
      Top             =   4005
      Width           =   795
   End
   Begin VB.TextBox txtIntervaloPerdidaStaminaLluvia 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   3975
      TabIndex        =   41
      Text            =   "0"
      Top             =   3330
      Width           =   900
   End
   Begin VB.TextBox txtCmdExec 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   3975
      TabIndex        =   40
      Text            =   "0"
      Top             =   4035
      Width           =   900
   End
   Begin VB.TextBox txtIntervaloVeneno 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   360
      TabIndex        =   33
      Text            =   "0"
      Top             =   3345
      Width           =   795
   End
   Begin VB.TextBox txtIntervaloParalizado 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   360
      TabIndex        =   32
      Text            =   "0"
      Top             =   4005
      Width           =   795
   End
   Begin VB.TextBox txtIntervaloInvisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1335
      TabIndex        =   31
      Text            =   "0"
      Top             =   3330
      Width           =   900
   End
   Begin VB.TextBox txtIntervaloInvocacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1335
      TabIndex        =   30
      Text            =   "0"
      Top             =   4035
      Width           =   900
   End
   Begin VB.TextBox txtIntervaloUserPuedeTrabajar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   6375
      TabIndex        =   25
      Text            =   "0"
      Top             =   1710
      Width           =   930
   End
   Begin VB.TextBox txtIntervaloParaConexion 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   6360
      TabIndex        =   24
      Text            =   "0"
      Top             =   1065
      Width           =   930
   End
   Begin VB.TextBox txtIntervaloSed 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4560
      TabIndex        =   20
      Text            =   "0"
      Top             =   1770
      Width           =   1410
   End
   Begin VB.TextBox txtIntervaloHambre 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Text            =   "0"
      Top             =   1095
      Width           =   1410
   End
   Begin VB.TextBox txtSanaIntervaloSinDescansar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3120
      TabIndex        =   16
      Text            =   "0"
      Top             =   1770
      Width           =   1050
   End
   Begin VB.TextBox txtSanaIntervaloDescansar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Text            =   "0"
      Top             =   1095
      Width           =   1050
   End
   Begin VB.TextBox txtStaminaIntervaloDescansar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1695
      TabIndex        =   12
      Text            =   "0"
      Top             =   1095
      Width           =   1050
   End
   Begin VB.TextBox txtStaminaIntervaloSinDescansar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "0"
      Top             =   1770
      Width           =   1050
   End
   Begin VB.TextBox txtIntervaloUserPuedeCastear 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   375
      TabIndex        =   8
      Text            =   "0"
      Top             =   1080
      Width           =   930
   End
   Begin VB.TextBox txtIntervaloUserPuedeAtacar 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Text            =   "0"
      Top             =   1755
      Width           =   930
   End
   Begin VB.CommandButton cmdAplicar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Aplicar"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   120
      X2              =   2160
      Y1              =   5760
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   0
      X2              =   2040
      Y1              =   4680
      Y2              =   5760
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "minutos."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1200
      TabIndex        =   59
      Top             =   5400
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WorldSave, cada :"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   240
      TabIndex        =   57
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Server:"
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
      Top             =   4680
      Width           =   765
   End
   Begin VB.Shape Shape14 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1905
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   4785
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puede Atacar"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5670
      TabIndex        =   53
      Top             =   3135
      Width           =   975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A.I."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5670
      TabIndex        =   52
      Top             =   3795
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "A.I."
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
      Left            =   5640
      TabIndex        =   49
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "NPC's"
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
      Left            =   5640
      TabIndex        =   48
      Top             =   2400
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wav Fx"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3030
      TabIndex        =   47
      Top             =   3135
      Width           =   555
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frio"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3030
      TabIndex        =   46
      Top             =   3795
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ener. Lluvia"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3975
      TabIndex        =   45
      Top             =   3120
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TimerExec"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3975
      TabIndex        =   44
      Top             =   3795
      Width           =   750
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Efecto cada..."
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
      Left            =   3000
      TabIndex        =   39
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Clima y Ambiente:"
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
      Left            =   3000
      TabIndex        =   38
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Veneno"
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   390
      TabIndex        =   37
      Top             =   3135
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paralizado"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   390
      TabIndex        =   36
      Top             =   3795
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1335
      TabIndex        =   35
      Top             =   3120
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invocacion"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1335
      TabIndex        =   34
      Top             =   3795
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Duración de..."
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
      TabIndex        =   29
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Magia:"
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
      TabIndex        =   28
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajo"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   6480
      TabIndex        =   27
      Top             =   1470
      Width           =   540
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IntervaloConex."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   6360
      TabIndex        =   26
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label34 
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
      Left            =   6480
      TabIndex        =   23
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sed"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4575
      TabIndex        =   22
      Top             =   1515
      Width           =   285
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hambre"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4590
      TabIndex        =   21
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sin descansar"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3135
      TabIndex        =   18
      Top             =   1515
      Width           =   1005
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descansando"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3150
      TabIndex        =   17
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descansando"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1710
      TabIndex        =   14
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sin descansar"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1695
      TabIndex        =   13
      Top             =   1515
      Width           =   1005
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lanza Spell"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   375
      TabIndex        =   10
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puede Atacar"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Hambre y Sed:"
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
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Sanación:"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Energia:"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Combate:"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label21 
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
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1260
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1740
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7545
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1980
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2265
   End
End
Attribute VB_Name = "frmG_C_Intervalos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAplicar_Click()
On Error Resume Next
Call AplicarIntervalos
End Sub

Private Sub cmdDefecto_Click()
On Error Resume Next
Call IntervalosPorDefecto
End Sub

Private Sub Command2_Click()
On Error GoTo Err

AplicarIntervalos
'Intervalos
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", str(SanaIntervaloSinDescansar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", str(StaminaIntervaloSinDescansar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", str(SanaIntervaloDescansar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", str(StaminaIntervaloDescansar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", str(IntervaloSed))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", str(IntervaloHambre))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", str(IntervaloVeneno))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", str(IntervaloParalizado))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", str(IntervaloInvisible))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", str(IntervaloFrio))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", str(IntervaloWavFx))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", str(IntervaloInvocacion))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", str(IntervaloParaConexion))

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", val(IntervaloUserPuedeCastear))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI", val(frmGeneral.TIMER_AI.Interval))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar", val(frmGeneral.NpcAtaca.Interval))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo", val(IntervaloUserPuedeTrabajar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", val(IntervaloUserPuedeAtacar))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia", val(frmGeneral.tLluvia.Interval))
Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec", val(frmGeneral.CmdExec.Interval))

Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS", val(MinutosWs))

Call FrmMensajes.msg("Nota", "Intervalos Guardados...")
Unload Me
Exit Sub
Err:
    Call FrmMensajes.msg("Error", "Error al intentar grabar los intervalos...")
End Sub

Private Sub Form_Load()
On Error Resume Next

Me.Left = 0
Me.Top = 0
Call IntervalosPorDefecto
End Sub

Sub IntervalosPorDefecto()
On Error Resume Next
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿ Intervalos del main loop ¿?¿?¿?¿?¿?¿?¿?¿?¿
txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
txtIntervaloSed.Text = IntervaloSed
txtIntervaloHambre.Text = IntervaloHambre
txtIntervaloVeneno.Text = IntervaloVeneno
txtIntervaloParalizado.Text = IntervaloParalizado
txtIntervaloInvisible.Text = IntervaloInvisible
txtIntervaloFrio.Text = IntervaloFrio
txtIntervaloWavFx.Text = IntervaloWavFx
txtIntervaloInvocacion.Text = IntervaloInvocacion
txtIntervaloParaConexion.Text = IntervaloParaConexion

'///////////////// TIMERS \\\\\\\\\\\\\\\\\\\

txtWS.Text = MinutosWs
If txtWS.Text = 0 Then
    Line1.Visible = True
    Line2.Visible = True
Else
    Line1.Visible = False
    Line2.Visible = False
End If
txtIntervaloUserPuedeCastear.Text = IntervaloUserPuedeCastear
txtNPCPuedeAtacar.Text = frmGeneral.NpcAtaca.Interval
txtAI.Text = frmGeneral.TIMER_AI.Interval
txtIntervaloUserPuedeTrabajar.Text = IntervaloUserPuedeTrabajar
txtIntervaloUserPuedeAtacar.Text = IntervaloUserPuedeAtacar
txtIntervaloPerdidaStaminaLluvia.Text = frmGeneral.tLluvia.Interval
txtCmdExec.Text = frmGeneral.CmdExec.Interval

End Sub

Public Sub AplicarIntervalos()
On Error Resume Next
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿ Intervalos del main loop ¿?¿?¿?¿?¿?¿?¿?¿?¿
SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
IntervaloSed = val(txtIntervaloSed.Text)
IntervaloHambre = val(txtIntervaloHambre.Text)
IntervaloVeneno = val(txtIntervaloVeneno.Text)
IntervaloParalizado = val(txtIntervaloParalizado.Text)
IntervaloInvisible = val(txtIntervaloInvisible.Text)
IntervaloFrio = val(txtIntervaloFrio.Text)
IntervaloWavFx = val(txtIntervaloWavFx.Text)
IntervaloInvocacion = val(txtIntervaloInvocacion.Text)
IntervaloParaConexion = val(txtIntervaloParaConexion.Text)

'///////////////// TIMERS \\\\\\\\\\\\\\\\\\\

IntervaloUserPuedeCastear = val(txtIntervaloUserPuedeCastear.Text)
frmGeneral.NpcAtaca.Interval = val(txtNPCPuedeAtacar.Text)
frmGeneral.TIMER_AI.Interval = val(txtAI.Text)
IntervaloUserPuedeTrabajar = val(txtIntervaloUserPuedeTrabajar.Text)
IntervaloUserPuedeAtacar = val(txtIntervaloUserPuedeAtacar.Text)
frmGeneral.tLluvia.Interval = val(txtIntervaloPerdidaStaminaLluvia.Text)
frmGeneral.CmdExec.Interval = val(txtCmdExec.Text)

MinutosWs = txtWS.Text

End Sub

Private Sub Text21_Change()

End Sub

Private Sub txtAI_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtCmdExec_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloFrio_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloHambre_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloInvisible_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloInvocacion_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloParaConexion_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloParalizado_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloPerdidaStaminaLluvia_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloSed_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloUserPuedeAtacar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloUserPuedeCastear_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloUserPuedeTrabajar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloVeneno_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtIntervaloWavFx_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtNPCPuedeAtacar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSanaIntervaloDescansar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSanaIntervaloSinDescansar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtStaminaIntervaloDescansar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtStaminaIntervaloSinDescansar_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtWS_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtWS_LostFocus()
On Error Resume Next
If IsNumeric(txtWS.Text) Then
    If txtWS.Text = 0 Then
        Line1.Visible = True
        Line2.Visible = True
    ElseIf txtWS.Text < 60 Then
        txtWS.Text = 60
        Line1.Visible = False
        Line2.Visible = False
    End If
Else
    txtWS.Text = MinutosWs
End If
End Sub
