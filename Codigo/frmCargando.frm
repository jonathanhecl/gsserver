VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCargando 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   5730
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481.986
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   377
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar Cargar 
      Height          =   330
      Left            =   330
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Ver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   12
      Left            =   150
      TabIndex        =   6
      Top             =   59
      Width           =   105
   End
   Begin VB.Label Ver2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   48
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2580
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando, por favor espere..."
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   4
      Top             =   4548
      Width           =   3945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando, por favor espere..."
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   3
      Left            =   870
      TabIndex        =   3
      Top             =   4560
      Width           =   3945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   2820
      TabIndex        =   2
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " aa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   4935
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Private Sub Form_Load()
On Error Resume Next
'Label1(1).Caption = Label1(1).Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
frmGeneral.Visible = False
Me.Picture = LoadPicture(App.Path & "\logo.jpg")
Ver.Caption = frmGeneral.Tag
Shape1.Width = Ver.Width + 30
Call Sombra
End Sub

Sub Sombra()

Ver2.Caption = Ver.Caption
Label1(1).Caption = Label1(3).Caption

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' [GS] Viejo systray :P
If Me.Visible = False Then ' El systray en MDIForma se fallaba entonces use Cargando como interprete :D
    ' 34.13334 El mouse pasa por encima
    ' 34.2 Primer click
    ' 34.4 Segundo click
    ' Like 34.33333 Doble click primero
    mx = X / Screen.TwipsPerPixelX
    If mx = 34.4 Then
        PopupMenu frmGeneral.mnuPopupMenu
    ElseIf mx Like 34.33333 Then
        frmGeneral.WindowState = vbNormal
        frmGeneral.Visible = True
        frmGeneral.QuitarSysTray
    End If
End If
' [/GS]

End Sub


