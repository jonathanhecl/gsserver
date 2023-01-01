VERSION 5.00
Begin VB.Form FrmMensajes 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MensajeX"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "FrmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5400
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
   End
   Begin VB.TextBox TXT 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmMensajes.frx":000C
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "FrmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Public Sub msg(ByVal Titulo As String, ByVal Texto As String)
Me.Caption = Titulo
TXT.Text = Texto
Me.Show
End Sub

Private Sub Timer1_Timer()
Me.SetFocus
End Sub
