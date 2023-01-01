VERSION 5.00
Begin VB.Form frmG_Versiones 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Versiones"
   ClientHeight    =   5910
   ClientLeft      =   4260
   ClientTop       =   2580
   ClientWidth     =   6495
   Icon            =   "frmG_Versiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   6495
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   5655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmG_Versiones.frx":1042
      Top             =   120
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   5685
      Left            =   110
      Top             =   110
      Width           =   6285
   End
End
Attribute VB_Name = "frmG_Versiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = 0
Me.Top = 0

End Sub
