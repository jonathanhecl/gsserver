VERSION 5.00
Begin VB.Form frmG_Configurar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Configuración"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmG_Configurar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   4650
   Begin VB.CommandButton cmdOpciones 
      BackColor       =   &H0000FF00&
      Caption         =   "Configurar &Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdIntervalos 
      BackColor       =   &H0000FF00&
      Caption         =   "Configurar &Invervalos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4380
   End
End
Attribute VB_Name = "frmG_Configurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()

End Sub

Private Sub cmdIntervalos_Click()
frmG_C_Intervalos.Show
frmG_C_Intervalos.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdOpciones_Click()
frmG_T_OPCIONES.Show
frmG_T_OPCIONES.SetFocus
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
End Sub
