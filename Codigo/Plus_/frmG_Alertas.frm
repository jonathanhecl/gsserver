VERSION 5.00
Begin VB.Form frmG_Alertas 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Alertas e Informaciones"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmG_Alertas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   7590
   Begin VB.CommandButton cmdLimpiar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Limpiar Reportes"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   7335
   End
   Begin VB.ListBox Al 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4905
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmG_Alertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpiar_Click()
Al.Clear
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Me.WindowState = 0
Me.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide
If frmGeneral.mnuCerrarCorrectamente.Checked = True Then Exit Sub
If frmGeneral.mnuCerrar.Checked = True Then Exit Sub
Cancel = True
End Sub
