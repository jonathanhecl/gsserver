VERSION 5.00
Begin VB.Form EDPJ_Borrado 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Escaneador de PJ's || Borrador de PJ's"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "EDPJ_Borrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox KillBAN 
      BackColor       =   &H00004000&
      Caption         =   "Borrar TODOS los Usuarios Baneados"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   3135
   End
   Begin VB.FileListBox archivos 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   2370
      Hidden          =   -1  'True
      Left            =   5520
      Pattern         =   "*.chr"
      TabIndex        =   9
      Top             =   5800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&BORRAR YA!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "30"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "15"
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00004000&
      Caption         =   "de Nivel menor a"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00004000&
      Caption         =   "de Nivel mayor a"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   2520
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "dias."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "...y que estén Abandonados por más o igual a"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "BORRADOR DE PJ's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1785
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5475
   End
End
Attribute VB_Name = "EDPJ_Borrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
archivos.Path = App.Path & "\Charfile\"

If IsNumeric(Text2.Text) = False Then
    MsgBox "La cantida de meses debe ser numerica."
    Exit Sub
End If
If IsNumeric(Text1.Text) = False Then
    MsgBox "El nivel debe ser numerico."
    Exit Sub
End If
If Text1 < 0 Or Text2 < 0 Then
    MsgBox "No son validos los numeros negativos."
    Exit Sub
End If

KillBAN.Enabled = False
Command1.Enabled = False
Command13.Enabled = False
Dim SINO As Boolean
SINO = (Option1.Value = True)
Option1.Enabled = False
Option2.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Dim i As Long
Dim nivel As Integer
Dim conta As Integer
DoEvents
archivos.Refresh
conta = 0
For i = 0 To archivos.ListCount - 1
    nivel = (GetVar(archivos.Path & "\" & archivos.List(i), "STATS", "ELV"))
    If nivel = 0 Then GoTo siguienteNW:
    If archivos.List(i) = "" Then GoTo siguienteNW:
    If SINO = True Then
        If nivel > val(Text1.Text) Then
            If (Date - GetFileInfo(archivos.Path & "\" & archivos.List(i))) >= Text2 Then
                If EsDios(ReadField(1, archivos.List(i), Asc("."))) Or EsSemiDios(ReadField(1, archivos.List(i), Asc("."))) Or EsConsejero(ReadField(1, archivos.List(i), Asc("."))) Then
                    ' No hago nada
                Else
                    Call MatarPersonaje(ReadField(1, archivos.List(i), Asc(".")))
                    conta = conta + 1
                    i = i - 1
                    archivos.Refresh
                End If
            End If
        End If
    Else
        If nivel < val(Text1.Text) Then
            If (Date - GetFileInfo(archivos.Path & "\" & archivos.List(i))) >= Text2 Then
                If EsDios(ReadField(1, archivos.List(i), Asc("."))) Or EsSemiDios(ReadField(1, archivos.List(i), Asc("."))) Or EsConsejero(ReadField(1, archivos.List(i), Asc("."))) Then
                    ' No hago nada
                Else
                    conta = conta + 1
                    Call MatarPersonaje(ReadField(1, archivos.List(i), Asc(".")))
                    i = i - 1
                    archivos.Refresh
                End If
            End If
        End If
    End If
    If KillBAN.Value = 1 Then
        If (GetVar(archivos.Path & "\" & archivos.List(i), "FLAGS", "BAN")) = 1 Then
            Call MatarPersonaje(ReadField(1, archivos.List(i), Asc(".")))
            conta = conta + 1
            i = i - 1
            archivos.Refresh
        End If
    End If
siguienteNW:
Next
If conta = 0 Then MsgBox "No se elimino ningun personaje."
If conta = 1 Then MsgBox "Solo un personaje fue eliminado."
If conta > 1 Then MsgBox "Fueron eliminados " & conta & " personajes."
Command1.Enabled = True
Command13.Enabled = True
Option1.Enabled = True
KillBAN.Enabled = True
Option2.Enabled = True
If SINO = True Then Option1.Value = True
Text1.Enabled = True
Text2.Enabled = True
End Sub

Private Sub Command13_Click()
On Error Resume Next
EscaneadorDePJs.Show
EscaneadorDePJs.archivos.Refresh
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Command13.Enabled = False Then Cancel = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub
