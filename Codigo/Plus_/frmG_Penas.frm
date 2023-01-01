VERSION 5.00
Begin VB.Form frmG_Penas 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Penas"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmG_Penas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   7125
   Begin VB.CommandButton cmdDefecto 
      BackColor       =   &H0000FF00&
      Caption         =   "&Desbanear"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1875
   End
   Begin VB.ListBox BanList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2955
      ItemData        =   "frmG_Penas.frx":1042
      Left            =   240
      List            =   "frmG_Penas.frx":104C
      TabIndex        =   1
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "&Baneados"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6915
   End
End
Attribute VB_Name = "frmG_Penas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
