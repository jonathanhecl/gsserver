VERSION 5.00
Begin VB.Form frmG_Trafico 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Trafico"
   ClientHeight    =   3225
   ClientLeft      =   5580
   ClientTop       =   4365
   ClientWidth     =   4695
   Icon            =   "frmG_Trafico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   Begin VB.ListBox LstTrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2955
      ItemData        =   "frmG_Trafico.frx":1042
      Left            =   120
      List            =   "frmG_Trafico.frx":1044
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3025
      Left            =   105
      Top             =   105
      Width           =   4500
   End
End
Attribute VB_Name = "frmG_Trafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
End Sub

