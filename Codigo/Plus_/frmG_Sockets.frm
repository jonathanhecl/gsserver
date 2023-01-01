VERSION 5.00
Begin VB.Form frmG_Sockets 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Socket's"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmG_Sockets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   9285
   Begin VB.Timer Actualizador 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   240
      Top             =   4440
   End
   Begin VB.TextBox Requests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2040
      Left            =   4935
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   810
      Width           =   3690
   End
   Begin VB.TextBox Errores 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   4935
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3180
      Width           =   3690
   End
   Begin VB.CheckBox chkDebug 
      BackColor       =   &H0000FF00&
      Caption         =   "&Debug Socket's > > >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   3855
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
      Left            =   270
      TabIndex        =   4
      Top             =   4500
      Width           =   4245
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
      Height          =   390
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   4290
   End
   Begin VB.CommandButton cmdLiberar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Liberar todos los Socket's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3945
      Width           =   4290
   End
   Begin VB.ListBox LstSocket 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   285
      TabIndex        =   2
      Top             =   420
      Width           =   4215
   End
   Begin VB.CommandButton cmdLimpiar 
      BackColor       =   &H0000FF00&
      Caption         =   "Limpiar Request's"
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
      TabIndex        =   12
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdReiniciar 
      BackColor       =   &H0000FF00&
      Caption         =   "Reiniciar todos los Socket's"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Request's:"
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   4965
      TabIndex        =   10
      Top             =   600
      Width           =   2685
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Errores:"
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Top             =   2955
      Width           =   2685
   End
   Begin VB.Label Jugando 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Debug Socket's:"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1725
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   5415
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4020
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2235
      Left            =   270
      Top             =   410
      Width           =   4260
   End
   Begin VB.Label Estado 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   285
      TabIndex        =   3
      Top             =   2880
      Width           =   4230
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   270
      Top             =   2800
      Width           =   4260
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   5415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4620
   End
End
Attribute VB_Name = "frmG_Sockets"
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

Private Sub chkDebug_Click()
If chkDebug.Value = 1 Then  ' Ver Debug
    Shape4.Visible = True
    Requests.Visible = True
    Errores.Visible = True
    cmdReiniciar.Visible = True
    cmdLimpiar.Visible = True
    Me.Width = 9045
    DebugSocket = True
Else                        ' Ocultar Debug
    Me.Width = 4950
    Shape4.Visible = False
    Requests.Visible = False
    Errores.Visible = False
    cmdReiniciar.Visible = False
    cmdLimpiar.Visible = False
    DebugSocket = False
End If
End Sub

Private Sub cmdActualizar_Click()

cmdActualizar.Enabled = False

LstSocket.Clear

Dim c As Integer
Dim i As Integer

For i = 1 To MaxUsers
    LstSocket.AddItem "Socket " & i & " - " & UserList(i).ConnID
    If UserList(i).ConnID <> -1 Then c = c + 1
Next i

If c = MaxUsers Then
    Estado.Caption = "No hay sockets vacios!"
Else
    Estado.Caption = "Hay " & MaxUsers - c & " sockets vacios!"
End If
cmdActualizar.Enabled = True

End Sub

Private Sub cmdLiberar_Click()
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
Next i

End Sub

Private Sub cmdLimpiar_Click()
frmG_Sockets.Requests.Text = ""
End Sub

Private Sub cmdReiniciar_Click()
Call ReloadSokcet
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = 4950
Shape4.Visible = False
Requests.Visible = False
Errores.Visible = False
cmdReiniciar.Visible = False
cmdLimpiar.Visible = False
Call cmdActualizar_Click
End Sub
