VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmG_Usuarios 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios"
   ClientHeight    =   4950
   ClientLeft      =   5370
   ClientTop       =   4125
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView LstUsuarios 
      Height          =   3135
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nick"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vida"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Mana"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Socket"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox Actualizar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      Caption         =   "Actualizar automaticamente"
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
      TabIndex        =   2
      Top             =   4320
      Width           =   5415
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Actualizar"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3230
      Left            =   320
      Top             =   520
      Width           =   5500
   End
   Begin VB.Label Jugando 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Usuarios Jugando: "
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
      TabIndex        =   0
      Top             =   240
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmG_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()

End Sub
