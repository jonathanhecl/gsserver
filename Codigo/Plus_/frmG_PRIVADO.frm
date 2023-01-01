VERSION 5.00
Begin VB.Form frmG_PRIVADO 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Privado"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "frmG_PRIVADO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TEXTO 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4370
      Width           =   6975
   End
   Begin VB.ListBox CHAT 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
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
      Height          =   4125
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cancelar"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   6930
   End
End
Attribute VB_Name = "frmG_PRIVADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()

End Sub

Private Sub cmdCancelar_Click()
PRIVADO_CON_EL_HOST = 0
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If PRIVADO_CON_EL_HOST <= 0 Then Exit Sub
Call SendData(ToIndex, PRIVADO_CON_EL_HOST, 0, "||EL HOST A CERRADO EL PRIVADO." & FONTTYPE_WARNING)
Call LogGM("PRIVADOS", "Fin privado con " & UserList(PRIVADO_CON_EL_HOST).Name, False)
Call cmdCancelar_Click
End Sub

Private Sub TEXTO_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub TEXTO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TEXTO = "" Then Exit Sub
    Call SendData(ToIndex, PRIVADO_CON_EL_HOST, 0, "||<Host> " & TEXTO.Text & FONTTYPE_WARNING & ENDC)
    Call MensPrivado("<Host> " & TEXTO.Text)
    TEXTO.Text = ""
End If
End Sub

Public Sub MensPrivado(ByVal Mensaje As String)
    CHAT.AddItem ReadField(1, Mensaje, Asc("~"))
    Call LogGM("PRIVADOS", ReadField(1, Mensaje, Asc("~")), False)
    If CHAT.ListCount > 71 Then CHAT.RemoveItem 0
    CHAT.ListIndex = CHAT.ListCount - 1
End Sub
