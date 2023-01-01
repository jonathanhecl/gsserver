VERSION 5.00
Begin VB.Form frmMantenimiento 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Mantenimiento"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3060
   Icon            =   "frmMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MAN 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Detener Conteo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Segundos."
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
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label SEG 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "El servidor se auto-ejecutara en..."
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Command1.Caption = "&Detener Conteo" Then
        Command1.Caption = "&Continuar con el Conteo"
        ' Detiene el conteo
        SEG.ForeColor = vbYellow
        MAN.Enabled = False
    Else
        Command1.Caption = "&Detener Conteo"
        ' Inicia el conteo
        SEG.ForeColor = vbGreen
        MAN.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Command1.Caption = "&Continuar con el Conteo"
' Detiene el conteo
SEG.ForeColor = vbYellow
MAN.Enabled = False
Dim Res
Res = MsgBox("Si cierra el sistema de Mantenimiento el servidor no se auto-ejecutara, ¿esta seguro que desea cerrarlo?", vbCritical + vbYesNo, "ALERTA")
' 0 cerrado con click en X
If Res = vbNo Then Cancel = True
End Sub

Private Sub MAN_Timer()
    SEG.Caption = SEG.Caption - 1
    If SEG.Caption < 1 Then
        Call Shell(App.Path & "\" & App.EXEName & ".exe -ejecutarigual", vbNormalFocus)
        End
    End If
End Sub
