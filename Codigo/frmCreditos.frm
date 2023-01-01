VERSION 5.00
Begin VB.Form frmCreditos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Acerca de..."
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1320
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4800
      Picture         =   "frmCreditos.frx":1042
      Stretch         =   -1  'True
      ToolTipText     =   "Argentum Online Server"
      Top             =   240
      Width           =   585
   End
   Begin VB.Label SS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   5
      Left            =   960
      TabIndex        =   12
      Top             =   3480
      Width           =   3105
   End
   Begin VB.Label SS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5040
      TabIndex        =   11
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label SS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5040
      TabIndex        =   10
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label SS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label SS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label SS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   840
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   3480
      Width           =   3105
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   4
      Left            =   2445
      TabIndex        =   5
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   3
      Left            =   2445
      TabIndex        =   4
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   2
      Left            =   2445
      TabIndex        =   3
      Top             =   2280
      Width           =   105
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   1
      Left            =   2445
      TabIndex        =   2
      Top             =   1920
      Width           =   105
   End
   Begin VB.Label TXT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   0
      Left            =   2445
      TabIndex        =   1
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Cerrar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   2250
      TabIndex        =   0
      Top             =   4990
      Width           =   675
   End
   Begin VB.Image xCerrar 
      Height          =   615
      Left            =   360
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Shape sCerrar 
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   360
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Shape Fondo 
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   360
      Top             =   240
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   4440
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Cerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sCerrar.FillColor = vbRed
End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sCerrar.FillColor <> &HC0& And Button <> 1 Then sCerrar.FillColor = &HC0&
End Sub

Private Sub Form_Load()
On Error Resume Next

Me.Picture = LoadPicture(App.Path & "\logo.jpg")
TXT(0).Caption = "GS Server AO " & frmGeneral.Tag
Me.Caption = "Acerca de " & TXT(0).Caption & "..."
TXT(1).Caption = "Programado por ^[GS]^"
TXT(2).Caption = "Web site: http://www.gs-zone.com.ar"
TXT(3).Caption = "E-mail: gshaxor@gmail.com"
TXT(4).Caption = "(r) NMS Optimized"
TXT(5).Caption = "Para ver los agradecimientos ejecutar /CREDITOS, en el juego. ;)"
For i = 0 To 5
    SS(i).ForeColor = vbWhite
    TXT(i).ForeColor = vbBlack
    SS(i).Caption = TXT(i).Caption
    SS(i).Left = TXT(i).Left - 10
    SS(i).Top = TXT(i).Top - 20
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sCerrar.FillColor <> &H80& Then sCerrar.FillColor = &H80&
End Sub

Private Sub Label1_Click()



End Sub

Private Sub Image2_Click()
On Error Resume Next
    Call Shell("explorer http://www.gs-zone.com.ar", vbMaximizedFocus)
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sCerrar.FillColor <> &H80& Then sCerrar.FillColor = &H80&
End Sub

Private Sub TXT_Click(Index As Integer)
On Error Resume Next
If Index = 2 Then
    Call Shell("explorer http://www.gs-zone.com.ar", vbMaximizedFocus)
ElseIf Index = 3 Then
    Clipboard.Clear
    Clipboard.SetText "gshaxor@gmail.com"
    Call FrmMensajes.msg("Nota", "E-mail copiado.")
End If
End Sub

Private Sub SS_Click(Index As Integer)
On Error Resume Next
If Index = 2 Then
    Call Shell("explorer http://www.gs-zone.com.ar", vbMaximizedFocus)
ElseIf Index = 3 Then
    Clipboard.Clear
    Clipboard.SetText "gshaxor@gmail.com"
    Call FrmMensajes.msg("Nota", "E-mail copiado.")
End If
End Sub

Private Sub xCerrar_Click()
Unload Me
End Sub

Private Sub xCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sCerrar.FillColor = vbRed
End Sub

Private Sub xCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sCerrar.FillColor <> &HC0& And Button <> 1 Then sCerrar.FillColor = &HC0&
End Sub
