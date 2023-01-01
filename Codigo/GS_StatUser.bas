Attribute VB_Name = "GS_StatUser"
' SendStatUser



Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status:" & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If

Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)
' [GS] Nuevas bobadas
Call SendData(ToIndex, sendIndex, 0, "||Dados: " & UserList(UserIndex).Stats.UserAtributos(1) & ", " & UserList(UserIndex).Stats.UserAtributos(2) & ", " & UserList(UserIndex).Stats.UserAtributos(3) & ", " & UserList(UserIndex).Stats.UserAtributos(4) & ", " & UserList(UserIndex).Stats.UserAtributos(5) & FONTTYPE_INFO)
' [GS]
' [GS] Estadisticas extra

LoopC = UserList(UserIndex).flags.TiempoOnline
Do
    If LoopC < 60 Then Exit Do
    LoopC = LoopC - 60
Loop
Call SendData(ToIndex, sendIndex, 0, "||Tiempo online: " & (UserList(UserIndex).flags.TiempoOnline - LoopC) / 60 & " hs con " & (LoopC) & " minutos. - PJs matados: " & (UserList(UserIndex).Stats.UsuariosMatados + UserList(UserIndex).Stats.CriminalesMatados) & FONTTYPE_INFO)
' [/GS]

If sendIndex <> UserIndex Then Call SendData(ToIndex, sendIndex, 0, "||E-mail: " & UserList(UserIndex).Email & FONTTYPE_INFO)

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)
Dim FileNamE As String
FileNamE = nombre & ".chr"
If FileExist(CharPath & FileNamE, vbNormal) = False Then
    Call SendData(ToIndex, sendIndex, 0, "||Pj Inexistente" & FONTTYPE_INFO)
Else

    Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & nombre & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & GetVar(CharPath & FileNamE, "stats", "elv") & "  EXP: " & GetVar(CharPath & FileNamE, "stats", "Exp") & "/" & GetVar(CharPath & FileNamE, "stats", "elu") & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & GetVar(CharPath & FileNamE, "stats", "minsta") & "/" & GetVar(CharPath & FileNamE, "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Salud: " & GetVar(CharPath & FileNamE, "stats", "MinHP") & "/" & GetVar(CharPath & FileNamE, "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & FileNamE, "Stats", "MinMAN") & "/" & GetVar(CharPath & FileNamE, "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & GetVar(CharPath & FileNamE, "stats", "MinHIT") & "/" & GetVar(CharPath & FileNamE, "stats", "MaxHIT") & FONTTYPE_INFO)
    
    Call SendData(ToIndex, sendIndex, 0, "||Oro: " & GetVar(CharPath & FileNamE, "stats", "GLD") & FONTTYPE_INFO)
    
    Call SendData(ToIndex, sendIndex, 0, "||E-mail: " & GetVar(CharPath & FileNamE, "CONTACTO", "Email") & FONTTYPE_INFO)
End If
Exit Sub

End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
    For j = 1 To MAX_INVENTORY_SLOTS
        Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
        End If
    Next
Else
    Call SendData(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call SendData(ToIndex, sendIndex, 0, "||Pj: " & .Name & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
'    Call SendData(ToIndex, sendIndex, 0, "||CriminalesMatados: " & .Faccion.CriminalesMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
'    Call SendData(ToIndex, sendIndex, 0, "||UsuariosMatados: " & .Stats.UsuariosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Clase: " & .clase & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim ban As String
Dim BanDetailPath As String

BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(ToIndex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
    ' 3 en uno :p
    Call SendData(ToIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
    ban = GetVar(CharFile, "FLAGS", "Ban")
    Call SendData(ToIndex, sendIndex, 0, "||Ban: " & ban & FONTTYPE_INFO)
    If ban = "1" Then
        Call SendData(ToIndex, sendIndex, 0, "||Ban por: " & GetVar(BanDetailPath, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
    End If
Else
    Call SendData(ToIndex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)
End If

End Sub
Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub
Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
        End If
    Next
Else
    Call SendData(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub



