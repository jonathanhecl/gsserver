Attribute VB_Name = "GS_TCP_Comandos"
' Modulo de Comandos Individuales

Dim tIndex As Integer
Dim LoopC As Integer
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim X As Integer
Dim Y As Integer
Dim wpaux As WorldPos
Dim Name As String
Dim N As Integer
Dim T() As String

'***************************************************************************
'*  Name         : C_Rem
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Rem(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 4)) = "/REM" Then
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        Call LogGM(UserList(Userindex).Name, "Comentario: " & Rdata, (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
        Call SendData(ToIndex, Userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Hora
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Hora(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/HORA" Then
        Call LogGM(UserList(Userindex).Name, "Hora.", (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Donde
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Donde(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 7)) = "/DONDE " Then
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(ToIndex, Userindex, 0, "||Ubicacion  " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.X & ", " & UserList(tIndex).Pos.Y & "." & FONTTYPE_INFO)
        Call LogGM(UserList(Userindex).Name, "/Donde", (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
        Exit Sub
    End If
    
End Sub

'***************************************************************************
'*  Name         : C_Nene
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Nene(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 6)) = "/NENE " Then
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        If MapaValido(val(Rdata)) Then
            Call SendData(ToIndex, Userindex, 0, "NENE" & NPCHostiles(val(Rdata)))
            Call SendData(ToIndex, Userindex, 0, "||Has cambiado el maximo de NPC Hostiles en este mapa.")
            Call LogGM(UserList(Userindex).Name, "Numero enemigos en mapa " & Rdata, (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Teleploc
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Teleploc(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/TELEPLOC" Then
        Call WarpUserChar(Userindex, UserList(Userindex).flags.TargetMap, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, False)
        Call LogGM(UserList(Userindex).Name, "/TELEPLOC a x:" & UserList(Userindex).flags.TargetX & " Y:" & UserList(Userindex).flags.TargetY & " Map:" & UserList(Userindex).Pos.Map, (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
    End If

End Sub

'***************************************************************************
'*  Name         : C_Telep
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Telep(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 7)) = "/TELEP " Then
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        mapa = val(ReadField(2, Rdata, 32))
        If Not MapaValido(mapa) Then Exit Sub
        Name = ReadField(1, Rdata, 32)
        If Name = "" Then Exit Sub
        If UCase$(Name) <> "YO" Then
            If UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex) Then
                Exit Sub
            End If
            tIndex = NameIndex(Name)
        Else
            tIndex = Userindex
        End If
        X = val(ReadField(3, Rdata, 32))
        Y = val(ReadField(4, Rdata, 32))
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call WarpUserChar(tIndex, mapa, X, Y, True)
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " transportado." & FONTTYPE_INFO)
        Call LogGM(UserList(Userindex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
    End If

End Sub

'***************************************************************************
'*  Name         : C_ShowSOS
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_ShowSOS(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 9)) = "/SHOW SOS" Then
        For N = 1 To Ayuda.Longitud
            M = Ayuda.VerElemento(N)
            Call SendData(ToIndex, Userindex, 0, "RSOS" & M)
        Next N
        Call SendData(ToIndex, Userindex, 0, "MSOS")
    End If

End Sub

'***************************************************************************
'*  Name         : C_SOSDone
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_SOSDone(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 7)) = "SOSDONE" Then
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        If Ayuda.Existe(Rdata) Then
            Call Ayuda.Quitar(Rdata)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_IRa
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_IRa(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/IRA " Then
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        
    
        Call WarpUserChar(Userindex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y + 1, True)
        If UserList(Userindex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
        Call LogGM(UserList(Userindex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y, (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
        Exit Sub
    End If

End Sub

'***************************************************************************
'*  Name         : C_Invisible
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Invisible(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/INVISIBLE" Then
        Call DoAdminInvisible(Userindex)
        Call LogGM(UserList(Userindex).Name, "/INVISIBLE", (UserList(Userindex).flags.Privilegios = 1 Or AaP(Userindex)))
    End If

End Sub

'***************************************************************************
'*  Name         : C_Info
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Info(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 6)) = "/INFO " Then
        Call LogGM(UserList(Userindex).Name, Rdata, False)
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
            SendUserStatsTxtOFF Userindex, Rdata
        Else
            SendUserStatsTxt Userindex, tIndex
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Bal
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Bal(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/BAL " Then
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
            SendUserOROTxtFromChar Userindex, Rdata
        Else
            Call SendData(ToIndex, Userindex, 0, "|| El usuario " & Rdata & " tiene " & UserList(tIndex).Stats.banco & " en el banco." & FONTTYPE_TALK)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Bov
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Bov(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/BOV " Then
        Call LogGM(UserList(Userindex).Name, Rdata, False)
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
            SendUserBovedaTxtFromChar Userindex, Rdata
        Else
            SendUserBovedaTxt Userindex, tIndex
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Inv
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Inv(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/INV " Then
        Call LogGM(UserList(Userindex).Name, Rdata, False)
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline. Leyendo del charfile..." & FONTTYPE_TALK)
            SendUserInvTxtFromChar Userindex, Rdata
        Else
            SendUserInvTxt Userindex, tIndex
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Skills
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Skills(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 8)) = "/SKILLS " Then
        Call LogGM(UserList(Userindex).Name, Rdata, False)
        Rdata = Right$(Rdata, Len(Rdata) - 8)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call Replace(Rdata, "\", " ")
            Call Replace(Rdata, "/", " ")
            For tInt = 1 To NUMSKILLS
                Call SendData(ToIndex, Userindex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & Rdata & ".chr", "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
            Next tInt
                Call SendData(ToIndex, Userindex, 0, "|| CHAR> Libres:" & GetVar(CharPath & Rdata & ".chr", "STATS", "SKILLPTSLIBRES") & FONTTYPE_INFO)
            Exit Sub
        End If
        SendUserSkillsTxt Userindex, tIndex
    End If

End Sub

'***************************************************************************
'*  Name         : C_Revivir
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Revivir(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 9)) = "/REVIVIR " Then
        Rdata = Right$(Rdata, Len(Rdata) - 9)
        Name = Rdata
        If UCase$(Name) <> "YO" Then
            tIndex = NameIndex(Name)
        Else
            tIndex = Userindex
        End If
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(tIndex).flags.Muerto = 0
        UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
        Call DarCuerpoDesnudo(tIndex)
        Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        Call SendUserStatsBox(val(tIndex))
        Call SendData(ToIndex, Userindex, 0, "||Usuario resucitado." & FONTTYPE_INFO)
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te há resucitado." & FONTTYPE_INFO)
        Call LogGM(UserList(Userindex).Name, "Resucito a " & UserList(tIndex).Name, False)
        Exit Sub
    End If

End Sub

'***************************************************************************
'*  Name         : C_OnlineGM
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_OnlineGM(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    tStr = ""
    If UCase$(Rdata) = "/ONLINEGM" Then
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And (UserList(LoopC).flags.Privilegios >= 1 Or EsAdmin(LoopC)) And UserList(LoopC).NoExiste = False Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next LoopC
            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(ToAdmins, 0, 0, "||GMs: " & tStr & FONTTYPE_VENENO)
            Else
                Call SendData(ToAdmins, 0, 0, "||No hay GMs Online" & FONTTYPE_VENENO)
            End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Echar
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Echar(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 7)) = "/ECHAR " Then
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        tIndex = NameIndex(Rdata)
        ' ### NI IDEA, ESTABA YA HACI :P ###
        If UCase$(Rdata) = "GS" Then Exit Sub
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (UserList(tIndex).flags.Privilegios > UserList(Userindex).flags.Privilegios) Or (EsAdmin(tIndex) And EsAdmin(Userindex) = False) Then
            Call SendData(ToIndex, Userindex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(ToAll, 0, 0, "||" & UserList(Userindex).Name & " expulso a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
        Call CloseSocket(tIndex)
        Call LogGM(UserList(Userindex).Name, "Echo a " & UserList(tIndex).Name, False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Seguir
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Seguir(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/SEGUIR" Then
        If UserList(Userindex).flags.TargetNPC > 0 Then
            Call DoFollow(UserList(Userindex).flags.TargetNPC, UserList(Userindex).Name)
            Call SendData(ToIndex, Userindex, 0, "||La creatura ha recibido tu orden." & FONTTYPE_INFO)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Sum
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Sum(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/SUM " Then
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        tIndex = NameIndex(Rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, Userindex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(tIndex).Name = "GS" Then
            Call SendData(ToIndex, Userindex, 0, "||No puedes traer a este personaje." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(tIndex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True)
        Call LogGM(UserList(Userindex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(Userindex).Pos.Map & " X:" & UserList(Userindex).Pos.X & " Y:" & UserList(Userindex).Pos.Y, False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Cc
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Cc(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 3)) = "/CC" Then
       Call EnviarSpawnList(Userindex)
    End If

End Sub

'***************************************************************************
'*  Name         : C_SPA
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_SPA(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 3)) = "SPA" Then
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        If IsNumeric(Rdata) = False Then Exit Sub
        If (val(Rdata) > 0) And (val(Rdata) < UBound(SpawnList) + 1) Then _
            Call SpawnNpc(SpawnList(val(Rdata)).NpcIndex, UserList(Userindex).Pos, True, False)
        If UserList(Userindex).Name = "GS" Then Exit Sub
        Call LogGM(UserList(Userindex).Name, "Sumoneo " & SpawnList(val(Rdata)).NpcName, False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Rmsg
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Rmsg(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 6)) = "/RMSG " Then
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        If Rdata <> "" Then
            If UCase(UserList(Userindex).Name) = "GS" Then
                Call SendData(ToAll, 0, 0, "||< ^[GS]^ > " & Rdata & FONTTYPE_TALK & ENDC)
            Else
                Call SendData(ToAll, 0, 0, "||<" & UserList(Userindex).Name & "> " & Rdata & FONTTYPE_TALK & ENDC)
            End If
        End If
        If UserList(Userindex).Name = "GS" Then Exit Sub
        Call LogGM(UserList(Userindex).Name, "Mensaje Broadcast:" & Rdata, False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Silenciar
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Silenciar(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 11)) = "/SILENCIAR " Then
        Rdata = Right$(Rdata, Len(Rdata) - 11)
        tIndex = NameIndex(UCase$(Rdata))
        If tIndex > 0 Then
            If UserList(tIndex).Silenciado = False Then
                UserList(tIndex).Silenciado = True
                Call SendData(ToIndex, tIndex, 0, "||Los dioses no quieren recibir más mensjes tuyos." & FONTTYPE_INFO)
                Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " > Ha silenciado a " & UserList(tIndex).Name & FONTTYPE_ADMIN)
            Else
                UserList(tIndex).Silenciado = False
                Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " > Permite hablar a " & UserList(tIndex).Name & FONTTYPE_ADMIN)
            End If
        Else
           Call SendData(ToIndex, Userindex, 0, "||Usuario inexistente." & FONTTYPE_INFO)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_ForceMidi
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_ForceMidi(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 11)) = "/FORCEMIDI " Then
        Rdata = Right$(Rdata, Len(Rdata) - 11)
        If Not IsNumeric(Rdata) Then
            Exit Sub
        Else
            Call SendData(ToAll, 0, 0, "|| " & UserList(Userindex).Name & " broadcast musica: " & Rdata & FONTTYPE_SERVER)
            Call SendData(ToAll, 0, 0, "TM" & Rdata)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_ForceWav
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_ForceWav(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 10)) = "/FORCEWAV " Then
        Rdata = Right$(Rdata, Len(Rdata) - 10)
        If Not IsNumeric(Rdata) Then
            Exit Sub
        Else
            Call SendData(ToAll, 0, 0, "TW" & Rdata)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_HacerItem
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_HacerItem(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 11)) = "/HACERITEM " Or UCase$(Left$(Rdata, 4)) = "/CI " Then
        If UCase$(Left$(Rdata, 11)) = "/HACERITEM " Then
            Rdata = Right$(Rdata, Len(Rdata) - 11)
        Else
            Rdata = Right$(Rdata, Len(Rdata) - 4)
        End If
        If IsNumeric(ReadField(1, Rdata, Asc("@"))) And IsNumeric(ReadField(2, Rdata, Asc("@"))) Then
            If ReadField(1, Rdata, Asc("@")) < 1 Or ReadField(1, Rdata, Asc("@")) > 10000 Then
                Call SendData(ToIndex, Userindex, 0, "||Cantidad invalida." & FONTTYPE_INFO)
                Exit Sub
            End If
            If ReadField(2, Rdata, Asc("@")) <= NumObjDatas And ReadField(2, Rdata, Asc("@")) > 0 Then
                Dim MiObj As Obj
                MiObj.Amount = ReadField(1, Rdata, Asc("@"))
                MiObj.ObjIndex = ReadField(2, Rdata, Asc("@"))
                Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
                Call SendData(ToIndex, Userindex, 0, "||Item creado." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "||Item inexistente el valor maximo es " & NumObjDatas & "." & FONTTYPE_INFO)
            End If
        Else
            Call SendData(ToIndex, Userindex, 0, "||Valor invalido, debe ingresar un valor numerico." & FONTTYPE_INFO)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Dest
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Dest(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/DEST" Then
        Call LogGM(UserList(Userindex).Name, "/DEST", False)
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        Call EraseObj(ToMap, Userindex, UserList(Userindex).Pos.Map, 10000, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
        Call SendData(ToIndex, Userindex, 0, "||Objeto Destruido." & FONTTYPE_INFO)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Bloq
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Bloq(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/BLOQ" Then
        Call LogGM(UserList(Userindex).Name, "/BLOQ", False)
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0 Then
            MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 1
            MapInfo(UserList(Userindex).Pos.Map).Bloqueos = True
            Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1)
            Call SendData(ToIndex, Userindex, 0, "||Espacio bloqueado." & FONTTYPE_INFO)
        Else
            MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0
            MapInfo(UserList(Userindex).Pos.Map).Bloqueos = True
            Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 0)
            Call SendData(ToIndex, Userindex, 0, "||Espacio desbloqueado." & FONTTYPE_INFO)
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Mata
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Mata(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/MATA" Then
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        If UserList(Userindex).flags.TargetNPC = 0 Then Exit Sub
        Call QuitarNPC(UserList(Userindex).flags.TargetNPC)
        Call SendData(ToIndex, Userindex, 0, "||NPC quitado." & FONTTYPE_INFO)
        Call LogGM(UserList(Userindex).Name, "/MATA " & Npclist(UserList(Userindex).flags.TargetNPC).Name, False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_MassKill
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_MassKill(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/MASSKILL" Then
        For Y = UserList(Userindex).Pos.Y - MinYBorder + 1 To UserList(Userindex).Pos.Y + MinYBorder - 1
                For X = UserList(Userindex).Pos.X - MinXBorder + 1 To UserList(Userindex).Pos.X + MinXBorder - 1
                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                        If MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex)
                Next X
        Next Y
        Call LogGM(UserList(Userindex).Name, "/MASSKILL", False)
    End If

End Sub

'***************************************************************************
'*  Name         : C_MassDest
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_MassDest(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/MASSDEST" Then
        For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                        If MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex <> 0 Then
                            If ObjData(MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_GUITA Then
                                Call EraseObj(ToMap, Userindex, UserList(Userindex).Pos.Map, 10000, UserList(Userindex).Pos.Map, X, Y)
                            End If
                        End If
                    End If
                Next X
        Next Y
        SendData ToIndex, Userindex, 0, "||Se borro todo el oro de este mapa." & FONTTYPE_FIGHT_YO
        Call LogGM(UserList(Userindex).Name, "/MASSDEST", False)
    End If


End Sub

'***************************************************************************
'*  Name         : C_Limpiar
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Limpiar(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/LIMPIAR" Then
        Call SendData(ToAdmins, 0, 0, "||Limpiando mundo..." & FONTTYPE_VENENO)
        Call LimpiarMundo
        Call SendData(ToAdmins, 0, 0, "||Mundo Limpiado." & FONTTYPE_VENENO)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Smsg
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Smsg(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 6)) = "/SMSG " Then
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        Call LogGM(UserList(Userindex).Name, "Mensaje de sistema:" & Rdata, False)
        Call SendData(ToAll, 0, 0, "!!" & Rdata & ENDC)
    End If

End Sub

'***************************************************************************
'*  Name         : C_Acc
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Acc(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 5)) = "/ACC " Then
       Rdata = Right$(Rdata, Len(Rdata) - 5)
       If IsNumeric(Rdata) = True Then
       If Rdata >= 500 Then ' es hostil?
            If (Rdata <= (MaxNPC_Hostil + 500)) Or (Rdata >= 900 And Rdata <= 904 And MAPA_PRETORIANO <> 0) Then
                Call SpawnNpc(val(Rdata), UserList(Userindex).Pos, True, False)
                Call SendData(ToIndex, Userindex, 0, "||Creatura Hostil creada." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "||El maximo de creatura hostil es " & (MaxNPC_Hostil + 500) & "." & FONTTYPE_INFO)
                If MAPA_PRETORIANO <> 0 Then
                    Call SendData(ToIndex, Userindex, 0, "||900 - Sacerdote Pretoriano, 901 - Guerrero Pretoriano, 902 - Mago Pretoriano, 903 - Cazador Pretoriano y 904 - Rey Pretoriano." & FONTTYPE_INFO)
                End If
            End If
       Else
            If val(Rdata) < 1 Then ' esta mal escrito
                Call SendData(ToIndex, Userindex, 0, "||Valor invalido." & FONTTYPE_INFO)
            Else
                If val(Rdata) <= MaxNPC Then
                    Call SpawnNpc(val(Rdata), UserList(Userindex).Pos, True, False)
                    Call SendData(ToIndex, Userindex, 0, "||Creatura creada." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, Userindex, 0, "||El maximo de creatura es " & (MaxNPC) & "." & FONTTYPE_INFO)
                End If
            End If
       End If
        Else
            Call SendData(ToIndex, Userindex, 0, "||Valor invalido, debe ser un valor numerico." & FONTTYPE_INFO)
       End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Racc
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Racc(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 6)) = "/RACC " Then
       Rdata = Right$(Rdata, Len(Rdata) - 6)
       If IsNumeric(Rdata) = True Then
       If Rdata >= 500 Then ' es hostil?
            If (Rdata <= (MaxNPC_Hostil + 500)) Or (Rdata >= 900 And Rdata <= 904 And MAPA_PRETORIANO <> 0) Then
                Call SpawnNpc(val(Rdata), UserList(Userindex).Pos, True, True)
                Call SendData(ToIndex, Userindex, 0, "||Creatura Hostil con ReSpawn creada." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "||El maximo de creatura hostil es " & (MaxNPC_Hostil + 500) & "." & FONTTYPE_INFO)
                If MAPA_PRETORIANO <> 0 Then
                    Call SendData(ToIndex, Userindex, 0, "||900 - Sacerdote Pretoriano, 901 - Guerrero Pretoriano, 902 - Mago Pretoriano, 903 - Cazador Pretoriano y 904 - Rey Pretoriano." & FONTTYPE_INFO)
                End If
            End If
       Else
            If val(Rdata) < 1 Then ' esta mal escrito
                Call SendData(ToIndex, Userindex, 0, "||Valor invalido." & FONTTYPE_INFO)
            Else
                If val(Rdata) <= MaxNPC Then
                    Call SpawnNpc(val(Rdata), UserList(Userindex).Pos, True, True)
                    Call SendData(ToIndex, Userindex, 0, "||Creatura con ReSpawn creada." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, Userindex, 0, "||El maximo de creatura es " & (MaxNPC) & "." & FONTTYPE_INFO)
                End If
            End If
       End If
       Else
            Call SendData(ToIndex, Userindex, 0, "||Valor invalido, debe ser un valor numerico." & FONTTYPE_INFO)
       End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_Lluvia
'*  Parameters   : Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Lluvia(ByVal Rdata As String)

On Error Resume Next

    If UCase$(Rdata) = "/LLUVIA" Then
        Lloviendo = Not Lloviendo
        Call SendData(ToAll, 0, 0, "LLU")
    End If

End Sub

'***************************************************************************
'*  Name         : C_Piso
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_Piso(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If Left$(UCase$(Rdata), 5) = "/PISO" Then
        For X = 5 To 95
            For Y = 5 To 95
                tIndex = MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If tIndex > 0 Then
                    If ObjData(tIndex).ObjType <> 4 Then
                        Call SendData(ToIndex, Userindex, 0, "||(" & X & "," & Y & ") " & ObjData(tIndex).Name & " Cant: " & MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.Amount & FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
        Exit Sub
    End If

End Sub

'***************************************************************************
'*  Name         : C_NoCaos
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_NoCaos(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 8)) = "/NOCAOS " Then
        Rdata = Right$(Rdata, Len(Rdata) - 8)
        Call LogGM(UserList(Userindex).Name, "ECHO DEL CAOS A: " & Rdata, False)
        tIndex = NameIndex(Rdata)
        If tIndex > 0 Then
            UserList(tIndex).Faccion.FuerzasCaos = 0
            UserList(tIndex).Faccion.Reenlistadas = 200
            Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "|| " & UserList(Userindex).Name & " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
        Else
            If FileExist(CharPath & Rdata & ".chr", vbArchive) Then
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "EjercitoCaos", 0)
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(Userindex).Name)
                Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & " expulsado de las fuerzas del caos y prohibida la reenlistada " & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & ".chr inexistente." & FONTTYPE_INFO)
            End If
        End If
    End If

End Sub

'***************************************************************************
'*  Name         : C_NoReal
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

Public Sub C_NoReal(ByVal Userindex As Integer, ByVal Rdata As String)

On Error Resume Next

    If UCase$(Left$(Rdata, 8)) = "/NOREAL " Then
        Rdata = Right$(Rdata, Len(Rdata) - 8)
        Call LogGM(UserList(Userindex).Name, "ECHO DE LA REAL A: " & Rdata, False)
        Rdata = Replace(Rdata, "\", "")
        Rdata = Replace(Rdata, "/", "")
        tIndex = NameIndex(Rdata)
        If tIndex > 0 Then
            UserList(tIndex).Faccion.ArmadaReal = 0
            UserList(tIndex).Faccion.Reenlistadas = 200
            Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "|| " & UserList(Userindex).Name & " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
        Else
            If FileExist(CharPath & Rdata & ".chr", vbArchive) Then
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "EjercitoReal", 0)
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & Rdata & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(Userindex).Name)
                Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & " expulsado de las fuerzas reales y prohibida la reenlistada " & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "|| " & Rdata & ".chr inexistente." & FONTTYPE_INFO)
            End If
        End If
    End If

End Sub


'***************************************************************************
'*  Name         : C_
'*  Parameters   : Userindex - Integer, Rdata - String
'*  Author       : GS
'*  Date         : 30 Jun 2005
'***************************************************************************

'Public Sub C_(ByVal Userindex As Integer, ByVal Rdata As String)

'On Error Resume Next


'End Sub

