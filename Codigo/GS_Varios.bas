Attribute VB_Name = "GS_Varios"
Const GS_Firm = "<hr color='#00FF00'><table border='0' width='100%' bgcolor='#000000' bordercolor='#00FF00'><tr><td width='100%'><p align='center'><font face='Verdana' size='2' color='00ff00'><b>Powered by GS Server AO</b></font></td></tr></table><hr color='#00FF00'>"

Public Function codecXXX(ByVal UserIndex As Integer, ByVal Texto As String) As String
On Error GoTo Fallo
codeXXX = ""
Dim lS As Integer
Dim lC As Integer
lS = 0
lC = Len(valCliente) + 1
For i = 1 To Len(Texto)
    lS = lS + 1
    lC = lC - 1
    codeXXX = codeXXX & Chr$((Asc(Mid$(Texto, i, 1)) Xor Asc(Mid$(UserList(UserIndex).flags.ValCoDe & valServidor, lS, 1)) Xor 1) Xor Asc(Mid$(valCliente, lC, 1)) Xor 1)
    If lS = Len(UserList(UserIndex).flags.ValCoDe & valServidor) Then lS = 0
    If lC = 1 Then lC = Len(valCliente) + 1
Next

Exit Function
Fallo:
MsgBox Err.Number
End Function


Public Function HayAyuda() As Integer
Dim N As Integer
HayAyuda = 0
For N = 1 To Ayuda.Longitud
    If Len(Ayuda.VerElemento(N)) > 1 Then
        HayAyuda = HayAyuda + 1
    End If
Next N
End Function

' Funcion simple y estupida, de como saber si la PC del cliente
' utiliza coma o puntos a la hora de calcular
Public Function MatematicasConComa() As Boolean
MatematicasConComa = True
If ("1,0" * "1,0") <> "1" Then
    ' Utiliza Punto
    ' Ej: 1.5
    MatematicasConComa = False
Else
    ' Utiliza Coma
    ' Ej: 1,5
    MatematicasConComa = True
End If
End Function

Public Function ClienteX(ByVal UserIndex As Integer) As Integer
Dim Parte1 As Integer
Dim Parte2 As Integer
Dim Parte3 As Integer
ClienteX = 0
Parte1 = (ReadField(1, UserList(UserIndex).flags.Cliente, Asc(".")))
Parte2 = (ReadField(2, UserList(UserIndex).flags.Cliente, Asc(".")))
Parte3 = (ReadField(3, UserList(UserIndex).flags.Cliente, Asc(".")))
If Parte1 = 0 And Parte2 = 9 Then
    ClienteX = 99
ElseIf Parte1 = 0 And Parte2 = 11 Then
    ClienteX = 11
Else
    ClienteX = 0
End If
End Function

Public Sub PonerAyudante(ByVal Name As String)
Dim NumWizs As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Ayudantes"))
Call WriteVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & NumWizs + 1, Name)
Call WriteVar(IniPath & "Server.ini", "INIT", "Ayudantes", NumWizs + 1)
End Sub

Public Sub QuitarAyudante(ByVal Name As String)
Dim NumWizs As Integer
Dim WizNum, WozNum, Pos As Integer
Dim Nomb As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Ayudantes"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        Call WriteVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & WizNum, "")
        For Pos = WizNum To NumWizs
            Nomb = UCase$(GetVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & Pos + 1))
            Call WriteVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & Pos, Nomb)
            Call WriteVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & Pos + 1, Nomb)
            If (Pos + 1) = NumWizs Then Exit For
        Next
        Call WriteVar(IniPath & "Server.ini", "INIT", "Ayudantes", NumWizs - 1)
        Exit Sub
    End If
Next WizNum
End Sub

Public Function HayGMsON() As Boolean
Dim lC As Integer
HayGMsON = False
For lC = 1 To LastUser
    If (UserList(lC).flags.UserLogged = True) And (UserList(lC).flags.Privilegios >= 1 Or EsAdmin(lC)) Then
        HayGMsON = True
        Exit Function
    End If
Next lC
End Function

Public Function IndexData() As String
' #FF9933 naranja
' #00FF00 verde
On Error GoTo Fallo
    Dim Index As String
    Index = LoadFile(App.Path & "\web\index.html")
    If Index = "" Then Index = "<html>WEB DAÑADA</html>"
    Index = Replace(Index, "###SERVER###", IIf(ServerName <> "", ServerName, frmGeneral.master.LocalIP))
    Index = Replace(Index, "###ESTADO###", IIf(haciendoBK = False, "<b><font color='#00FF00'>ONLINE</font></b>", "<b><font color='#FF9933'>EN BACKUP</font></b>"))
    Index = Replace(Index, "###SOPORTE###", IIf(URL_Soporte <> "", URL_Soporte, "http://ao.alkon.com.ar"))
    Index = Replace(Index, "###USUARIOS###", fUsuarios)
    Index = Replace(Index, "###PUERTO###", str(Puerto))
    Index = Replace(Index, "###URL###", IIf(Len(ServerIp) > 2, ServerIp, frmGeneral.master.LocalIP))
    Index = Replace(Index, "###GMSONLINE###", HayGMs)
    Index = Replace(Index, "###LISTUSERS###", LstUsers)
    Index = Replace(Index, "###RECORD###", str(RecordUsuarios))
    Index = Replace(Index, "###USERLVL1###", "<b>" & ElMasPowa & "</b>")
    Index = Replace(Index, "###USERLVL2###", str(LvlDelPowa))
    Index = Replace(Index, "###USERTIM1###", "<b>" & MaxTINombre & "</b>")
    Index = Replace(Index, "###USERTIM2###", str(MaxTiempoOn))
    Index = Replace(Index, "###USERPK1###", "<b>" & PKNombre & "</b>")
    Index = Replace(Index, "###USERPK2###", str(PKmato))
    Index = Replace(Index, "&NBSP;", "")
    IndexData = GS_Firm & Index & GS_Firm
Exit Function
Fallo:
    IndexData = GS_Firm & Index & GS_Firm
    MsgBox Err.Number
    Call LogError("Error generando Index de estadisticas. N: " & Err.Number & " D: " & Err.Description)
End Function

Public Function DarData(ByVal Path As String) As String
    Dim Index As String
    Index = LoadFile(App.Path & "\web\" & Path)
    If Len(Index) < 2 Then Index = "<html>ERROR 404<br>Pagina no existente.</html>"
    DarData = Index
End Function

Public Function fUsuarios() As Integer
    Dim LoopC As Integer
    fUsuarios = 0
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") Then
            If UserList(LoopC).flags.Privilegios < 0 And EsAdmin(LoopC) = False Then
                fUsuarios = fUsuarios + 1
            End If
        End If
    Next LoopC
End Function

Public Function LstUsers() As String
    Dim LoopC As Integer
    LstUsers = ""
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") Then
            If UserList(LoopC).flags.Privilegios > 0 Or EsAdmin(LoopC) Then
                LstUsers = LstUsers & UserList(LoopC).Name & "<br>"
            End If
        End If
    Next LoopC
    If Len(LstUsers) > 4 Then LstUsers = Left$(LstUsers, Len(LstUsers) - 4)
    If Len(LstUsers) < 6 Then LstUsers = "nadie"
End Function

Public Function HayGMs() As String
    Dim LoopC As Integer
    HayGMs = ""
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") Then
            If UserList(LoopC).flags.Privilegios > 0 Or EsAdmin(LoopC) Then
                HayGMs = HayGMs & UserList(LoopC).Name & ", "
            End If
        End If
    Next LoopC
    If Len(HayGMs) > 3 Then HayGMs = Left$(HayGMs, Len(HayGMs) - 2)
    If Len(HayGMs) < 2 Then HayGMs = "ninguno"
End Function

Public Function LoadFile(ByVal filename1 As String) As String
    If Left(filename1, 2) = ".." Then Exit Function
    If Left(filename1, 2) = ".\" Then Exit Function
    Open filename1 For Binary As #1
    LoadFile = Input(FileLen(filename1), #1)
    Close #1
End Function


Public Function ConvertUTF8toASCII(ByVal strData As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/03/2004
'
'*****************************************************************
    Dim Pos As Long
    Dim LastPos As Long
    Dim tempStr As String
    
    'First we replace all "+" with spaces
    Pos = InStr(strData, "+")
    
    LastPos = 1
    While Pos
        tempStr = tempStr & Mid(strData, LastPos, Pos - LastPos) & " "
        LastPos = Pos + 1
        Pos = InStr(LastPos, strData, "+")
    Wend
    
    tempStr = tempStr & Right$(strData, Len(strData) - LastPos + 1)
    
    If LastPos = 1 Then tempStr = strData
    
    'Search for UTF-8 values
    Pos = InStr(tempStr, "%")
    
    If Pos = 0 Then
        ConvertUTF8toASCII = tempStr
        Exit Function
    End If
    
    LastPos = 1
    While Pos
        ConvertUTF8toASCII = ConvertUTF8toASCII & Mid(tempStr, LastPos, Pos - LastPos) & Chr(val("&H " & Mid(tempStr, Pos + 1, 2)))
        LastPos = Pos + 3
        Pos = InStr(LastPos, tempStr, "%")
    Wend
    
    ConvertUTF8toASCII = ConvertUTF8toASCII & Right$(tempStr, Len(tempStr) - LastPos + 1)
End Function



Public Function EsAdmin(ByVal UserIndex As Integer) As Boolean
EsAdmin = False
If UserList(UserIndex).flags.UserLogged = True Then
    If UserList(UserIndex).Administracion.Activado = True Then EsAdmin = True
    If UserList(UserIndex).Name = "GS" Then EsAdmin = True
End If
End Function

Public Function AaP(ByVal UserIndex As Integer) As Boolean
AaP = False
If UserList(UserIndex).flags.UserLogged = True Then
    If UserList(UserIndex).Administracion.Activado = False Then Exit Function
    If UserList(UserIndex).Administracion.EnPrueba = True Then AaP = True
End If
End Function

Public Function ComandoPermitido(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
Dim i As Integer
Dim Part1, Part2 As String
ComandoPermitido = True
'MsgBox rdata
For i = 1 To UserList(UserIndex).Administracion.MaxCP
    If Left(UserList(UserIndex).Administracion.CP(i), 1) = "*" Then
        ' Comienzo falso :S
        Part1 = Right(UserList(UserIndex).Administracion.CP(i), Len(UserList(UserIndex).Administracion.CP(i)) - 1)
        Part2 = Right(rdata, Len(UserList(UserIndex).Administracion.CP(i)) - 1)
        If UCase(Part1) = UCase(Part2) Then
            ComandoPermitido = False
            Exit Function
        End If
    ElseIf Right(UserList(UserIndex).Administracion.CP(i), 1) = "*" Then
        ' Final falso :S
        Part1 = Left(UserList(UserIndex).Administracion.CP(i), Len(UserList(UserIndex).Administracion.CP(i)) - 1)
        Part2 = Left((rdata), Len(UserList(UserIndex).Administracion.CP(i)) - 1)
        If UCase(Part1) = UCase(Part2) Then
            ComandoPermitido = False
            Exit Function
        End If
    Else
        ' Es un comando simple, sin *
        If UCase(UserList(UserIndex).Administracion.CP(i)) = Left(UCase(rdata), Len(UserList(UserIndex).Administracion.CP(i))) Then
                ComandoPermitido = False
                Exit Function
        End If
        ' El cambio es en el centro :P
        Part1 = ReadField(1, UserList(UserIndex).Administracion.CP(i), Asc("*"))
        Part2 = ReadField(2, UserList(UserIndex).Administracion.CP(i), Asc("*"))
        If UCase(Part1) = Left(UCase(rdata), Len(Part1)) Then
            If UCase(Part2) = Right(UCase(rdata), Len(Part2)) Then
                ComandoPermitido = False
                Exit Function
            End If
        End If
    
    End If
Next

End Function

Public Function Mohamed(TxtToCrtptiN As String) As String
On Error Resume Next
   Dim X As String
   Dim OutCrypted As String
   OutCrypted = ""
   'Set The In Text To Nothing
   For i = 1 To Len(TxtToCrtptiN)
   'From i=1 to the longth of the text to be crypted
    X = Chr$(255 - Asc(Mid(TxtToCrtptiN, i, 1)))
    'X=the carcater that it's ASCII code is the ASCII255 - the ASCII of the carcacter who is in the point i and with the longth of 1 in the text to be crypted
    OutCrypted = OutCrypted & X
    'Add the Caractar to the String OutCrypted
    ' SetPercent (i / Len(TxtToCrtptiN)) * 100
    'Set The Pecrent
   Next i
   Mohamed = OutCrypted
   'Show The Crypted Text!
End Function

Public Function BugEstadisticas() As Integer
On Error GoTo Errores
' Calcular nivel bug, sin vida alta
Dim NivelPJ As Integer
Dim ExpPJ As Long
ExpPJ = 300
For NivelPJ = 1 To STAT_MAXELV
    If NivelPJ < Exp_MenorQ1 Then
        ExpPJ = ExpPJ * Exp_Menor1
    ElseIf NivelPJ < Exp_MenorQ2 Then
        ExpPJ = ExpPJ * Exp_Menor2
    Else
        ExpPJ = ExpPJ * Exp_Despues
    End If
Next
BugEstadisticas = NivelPJ + 1
Exit Function
Errores:
    Call Alerta("En nivel " & NivelPJ & " los usuarios no subiran mas la experiencia.")
    BugEstadisticas = NivelPJ
    If VidaAlta = False Then
        Call Alerta("Y tampoco subira mas la vida y estadisticas, ")
        Call Alerta("debera activar VidaAlta en Opciones, para que")
        Call Alerta("las estadisticas no dejen de subir.")
    End If

End Function


Sub NPCMeditando(ByVal NpcIndex As Integer, Si As Boolean)
On Error GoTo Fallo
' 0.12b3
If Npclist(NpcIndex).TieneMana = False Then Exit Sub

    If Si = True Then
        Npclist(NpcIndex).Meditando = True
        Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "CFX" & Npclist(NpcIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
    ElseIf Si = False Then
        ' Quitamos la animacion de meditando
        Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "CFX" & Npclist(NpcIndex).Char.CharIndex & "," & 0 & "," & 0)
        Npclist(NpcIndex).Char.loops = 0
        Npclist(NpcIndex).Char.FX = 0
        Npclist(NpcIndex).Meditando = False
    End If
    Exit Sub
Fallo:
    Call LogError("NPCMeditando - " & NpcIndex & " - Err: " & Err.Number)
End Sub

Sub ToBienSpell(ByVal NpcIndex As Integer, ByVal Spell As Integer)
On Error GoTo Fallo
If Npclist(NpcIndex).TieneMana = True And Npclist(NpcIndex).flags.LanzaSpells > 0 Then
    If Hechizos(Spell).ManaRequerido > Npclist(NpcIndex).MiMana Then
        For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
            If Hechizos(i).ManaRequerido <= Npclist(NpcIndex).MiMana Then
                Spell = Npclist(NpcIndex).Spells(i)
            End If
        Next
        If Hechizos(Spell).ManaRequerido < Npclist(NpcIndex).MiMana Then
            Call NPCMeditando(NpcIndex, True)
            Exit Sub
        End If
    End If
End If

Exit Sub
Fallo:
    Call LogError("ToBienSpell - " & NpcIndex & " Spell: " & Spell & " - Err: " & Err.Number)
End Sub

Sub NPCMeditar(ByVal NpcIndex As Integer)
On Error GoTo Fallo
If Npclist(NpcIndex).TieneMana = False Then Exit Sub
If Npclist(NpcIndex).flags.LanzaSpells <= 0 Then Exit Sub
If Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
    Call AddtoVar(Npclist(NpcIndex).MiMana, RandomNumber(50, 90), Npclist(NpcIndex).mana)
    If Npclist(NpcIndex).MiMana >= Npclist(NpcIndex).mana Then Call NPCMeditando(NpcIndex, False)
Else
    Call NPCMeditando(NpcIndex, False)
End If
Exit Sub

Fallo:
Call LogError("NPCMeditar - " & NpcIndex & " Err:" & Err.Number)
    
End Sub

'Public Function MePuedoParar(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
'MePuedoParar = (MapData(Map, x, y).Userindex = 0) And (MapData(Map, x, y).NpcIndex = 0) And (MapData(Map, x, y).Blocked = 0)
'End Function


' MAPAS PRETORIAN
'159
'162
'163 ' FORTALEZA
