VERSION 5.00
Begin VB.Form frmG_NPCDebug 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> NPC Debug"
   ClientHeight    =   3255
   ClientLeft      =   5370
   ClientTop       =   4575
   ClientWidth     =   4965
   Icon            =   "frmG_NPCDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   4965
   Begin VB.Timer Actualizador 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   240
      Top             =   2640
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
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   4455
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
      Height          =   420
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   " NPC's Activos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   " NPC's Libres:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   255
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   " Indice del Ultimo NPC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   " Maximo de NPC's:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4740
   End
End
Attribute VB_Name = "frmG_NPCDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizador_Timer()
On Error Resume Next
If cmdActualizar.Enabled <> False Then Call cmdActualizar_Click
End Sub

Private Sub Actualizar_Click()
If Actualizar.Value = 1 Then
    Actualizador.Enabled = True
Else
    Actualizador.Enabled = False
End If
End Sub

Private Sub cmdActualizar_Click()
cmdActualizar.Enabled = False

Dim i As Integer, k As Integer

For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive Then k = k + 1
Next i

Label1.Caption = " NPC's Activos: " & k
Label2.Caption = " NPC's Libres: " & MAXNPCS - k
Label3.Caption = " Indice del Ultimo NPC: " & LastNPC
Label4.Caption = " Maximo de NPC's: " & MAXNPCS

cmdActualizar.Enabled = True
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Call cmdActualizar_Click
End Sub
