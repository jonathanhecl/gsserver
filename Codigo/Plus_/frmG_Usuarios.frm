VERSION 5.00
Begin VB.Form frmG_Usuarios 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   5670
   ClientLeft      =   5370
   ClientTop       =   4170
   ClientWidth     =   5070
   Icon            =   "frmG_Usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   5070
End
Attribute VB_Name = "frmG_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizador_Timer()
On Error Resume Next
If cmdActualizar.Enabled = True Then Call cmdActualizar_Click
End Sub

Private Sub Actualizar_Click()
If Actualizar.Value = 1 Then
    Actualizador.Enabled = True
Else
    Actualizador.Enabled = False
End If
End Sub

Private Sub cmdActualizar_Click()
If frmG_Usuarios.Visible = False Then Exit Sub
cmdActualizar.Enabled = False
DoEvents
Dim numeroJuaZ As Integer
numeroJuaZ = 0
' La lista
Dim h As Long
LstUsuarios.Clear
For h = 1 To LastUser
    If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
        ' Nombre " - Nivel: " El nivel - Num:~" Numero
        LstUsuarios.AddItem UserList(h).Name & " - Nivel: " & UserList(h).Stats.ELV & " - Num:~" & str(h)
        numeroJuaZ = numeroJuaZ + 1
    ElseIf UserList(h).ConnID <> -1 Then
        LstUsuarios.AddItem "Iniciando usuario..."
    End If
Next h

NumUsers = numeroJuaZ

Jugando.Caption = "Usuarios Jugando: " & NumUsers

cmdActualizar.Enabled = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
EscaneadorDePJs.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim nick As String
Dim tIndex As Integer
If Len(NickDat.Caption) > 1 Then
    nick = ReadField(1, NickDat.Caption, Asc("-"))
    nick = Left(nick, Len(nick) - 1)
    tIndex = NameIndex(nick)
    If tIndex <= 0 Then Exit Sub
    Call SendData(ToIndex, tIndex, 0, "||HOST le dice a Usted: " & Mensaje.Text & FONTTYPE_WHISPER)
    Call LogCOSAS("Host", Time & " HOST le dice a Usted: " & Mensaje.Text)
    Mensaje.Text = ""
End If
End Sub

Private Sub Form_Activate()
Me.Show
Me.Visible = True
Call cmdActualizar_Click
Actualizar.Value = 1
Actualizador.Enabled = True
End Sub

Private Sub Form_Load()
If frmG_Main.Visible = True Then
    Me.Left = frmG_Main.Left + frmG_Main.Width
    Me.Top = 0
Else
    Me.Left = 0
    Me.Top = 0
End If
Me.Show
Me.Visible = True
Call cmdActualizar_Click
End Sub

Private Sub LstUsuarios_Click()
On Error Resume Next
NickDat.Caption = LstUsuarios.Text
End Sub

Private Sub LstUsuarios_DblClick()
'On Error Resume Next
'Dim nick As String
'Dim tIndex As Integer
'nick = ReadField(1, LstUsuarios.Text, Asc("-"))
'If nick = "" Then Exit Sub
'Nick = Left(nick, Len(nick) - 1)
'tIndex = NameIndex(nick)
'Call PJ.Vigilar(tIndex)
End Sub

Private Sub LstUsuarios_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
'If Button = 3 Or Button = 2 Then
'    Me.PopupMenu mnuOpc
'End If
End Sub

Private Sub mnuBander_Click()
On Error Resume Next
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UserList(NumIndex).ConnID < 0 Then
        Call FrmMensajes.MSG("Alerta", "No se encuentra online.")
        Exit Sub
    End If
    If Not InMapBounds(Banderbill.Map, Banderbill.x, Banderbill.y) Then Exit Sub
    If NumIndex < 0 Then Exit Sub
    Call WarpUserChar(NumIndex, Banderbill.Map, Banderbill.x, Banderbill.y, True)
    Call LogGM("HOST", "Transporto a " & UserList(NumIndex).Name, False)
End If
End Sub

Private Sub mnuCharlaenprivado_Click()
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UserList(NumIndex).ConnID < 0 Then
        MsgBox "No se encuentra online."
        Exit Sub
    End If
    If PRIVADO_CON_EL_HOST > 0 Then
        If frmG_PRIVADO.Visible = True Then
              Unload frmG_PRIVADO
        End If
        PRIVADO_CON_EL_HOST = 0
    End If
    If NumIndex < 0 Then Exit Sub
    Call SendData(ToIndex, NumIndex, 0, "||ESTAS EN UN SECCION PRIVADA CON EL HOST, SI DESEAS CANCELARLA ESCRIBE /CANCELAR." & FONTTYPE_WARNING)
    PRIVADO_CON_EL_HOST = NumIndex
    Call LogGM("PRIVADOS", "Inicio privado con " & UserList(PRIVADO_CON_EL_HOST).Name, False)
    frmG_PRIVADO.Show
End If
End Sub

Private Sub mnuExpulsar_Click()
On Error Resume Next
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UCase$(UserList(NumIndex).Name) = "GS" Then Exit Sub
    If UserList(NumIndex).ConnID < 0 Then
        MsgBox "No se encuentra online."
        Exit Sub
    End If
    If NumIndex < 0 Then Exit Sub
    Call SendData(ToAll, 0, 0, "||El <Host> expulso a " & UserList(NumIndex).Name & "." & FONTTYPE_INFO)
    Call CloseSocket(NumIndex)
    Call LogGM("HOST", "Echo a " & UserList(NumIndex).Name, False)
End If
End Sub

Private Sub mnuLindos_Click()
On Error Resume Next
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UserList(NumIndex).ConnID < 0 Then
        MsgBox "No se encuentra online."
        Exit Sub
    End If
    If Not InMapBounds(Lindos.Map, Lindos.x, Lindos.y) Then Exit Sub
    If NumIndex < 0 Then Exit Sub
    Call WarpUserChar(NumIndex, Lindos.Map, Lindos.x, Lindos.y, True)
    Call LogGM("HOST", "Transporto a " & UserList(NumIndex).Name, False)
End If
End Sub

Private Sub mnuNix_Click()
On Error Resume Next
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UserList(NumIndex).ConnID < 0 Then
        MsgBox "No se encuentra online."
        Exit Sub
    End If
    If Not InMapBounds(Nix.Map, Nix.x, Nix.y) Then Exit Sub
    If NumIndex < 0 Then Exit Sub
    Call WarpUserChar(NumIndex, Nix.Map, Nix.x, Nix.y, True)
    Call LogGM("HOST", "Transporto a " & UserList(NumIndex).Name, False)
End If
End Sub

Private Sub mnuUlla_Click()
On Error Resume Next
If LstUsuarios.ListIndex > -1 Then
    Dim NumIndex As Integer
    NumIndex = val(ReadField(2, LstUsuarios.ListIndex, Asc("~")))
    If UserList(NumIndex).ConnID < 0 Then
        MsgBox "No se encuentra online."
        Exit Sub
    End If
    If Not InMapBounds(Ullathorpe.Map, Ullathorpe.x, Ullathorpe.y) Then Exit Sub
    If NumIndex < 0 Then Exit Sub
    Call WarpUserChar(NumIndex, Ullathorpe.Map, Ullathorpe.x, Ullathorpe.y, True)
    Call LogGM("HOST", "Transporto a " & UserList(NumIndex).Name, False)
End If

End Sub

