Attribute VB_Name = "GS_Ayuda"
'Type tComando
'    Comando As String
'    Categoria As Byte
'    Aprendiz As Boolean
'    Admin As Boolean
'    Privilegios As Byte
'    Syntaxis As String
'    Descripcion As String
'End Type
'Public MaxComandos As Integer
'Public Comando() As tComando



'Public Sub CargarAyuda()
'On Error Resume Next
'Dim Num As Integer
''Set Leer = LeerAy
'Dim LeerAyuda As clsLeerInis
'LeerAyuda.Abrir App.Path & "\Comandos.dat"
'MaxComandos = val(LeerAyuda.DarValor("INIT", "NumComandos"))'
'
'frmCargando.Cargar.Value = 0
'frmCargando.Cargar.Min = 0
'frmCargando.Cargar.max = MaxComandos
'frmGeneral.ProG1.Min = 0
'frmGeneral.ProG1.max = MaxComandos
'frmGeneral.ProG1.Value = 0
'frmGeneral.ProG1.Visible = True'

'ReDim Preserve Comando(1 To MaxComandos) As tComando

'For Num = 1 To MaxComandos
'    Comando(Num).Comando = LeerAyuda.DarValor("COMANDO" & i, "Comando")
'    Comando(Num).Categoria = val(LeerAyuda.DarValor("COMANDO" & i, "Categoria"))
'    Comando(Num).Syntaxis = Leer.DarValor("COMANDO" & i, "Syntaxis")
'    If Comando(Num).Syntaxis = "" Then Comando(Num).Syntaxis = Comando(Num).Comando
'    Comando(Num).Descripcion = Leer.DarValor("COMANDO" & i, "Descripcion")
'    Comando(Num).Privilegios = val(Leer.DarValor("COMANDO" & i, "Privilegios"))
'    If Comando(Num).Privilegios < 1 Then Comando(Num).Privilegios = 1
'    If Comando(Num).Privilegios < 3 Then
'        Comando(Num).Aprendiz = True
'    Else
'        Comando(Num).Aprendiz = False
'    End If
'    If Comando(Num).Privilegios > 1 Then
'        Comando(Num).Admin = True
'    Else
'        Comando(Num).Admin = False
'    End If
'    If frmCargando.Visible Then
'        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
'        frmCargando.Label1(0).Caption = Comando(Num).Comando
'    End If
'    If frmGeneral.Visible Then frmGeneral.ProG1.Value = frmGeneral.ProG1.Value + 1
'
'Next
'frmGeneral.ProG1.Visible = False
'Exit Sub
'Fallo:
'MsgBox Err.Number & " - " & Err.Description
'End Sub


Sub DarAyuda(ByVal Userindex As Integer, ByVal rdata As String)
    'Call AyudaLoad
    'Call CargarAyuda
    'Dim i, j As Integer
    'Dim Tempo As String
    'For i = 1 To MaxComandos
    '    If UCase$(Comando(i).Comando) = UCase$(Right(rdata, Len(rdata) - 7)) Then
    '        Call SendData(ToIndex, Userindex, 0, "||Syntaxis: " & Comando(i).Comando & FONTTYPE_VENENO)
    '        Tempo = ""
    '        For j = 1 To Len(Comando(i).Descripcion)
    '            If Mid(Comando(i).Descripcion, j, 1) = "|" Then
    '                Call SendData(ToIndex, Userindex, 0, "||" & Tempo & FONTTYPE_INFX)
    '                Tempo = ""
    '            Else
    '                Tempo = Tempo & Mid(Comando(i).Descripcion, j, 1)
    '            End If
    '        Next
    '        If Len(Tempo) > 0 Then
    '            Call SendData(ToIndex, Userindex, 0, "||" & Tempo & FONTTYPE_INFX)
    '            Tempo = ""
    '        End If
    '        Exit Sub
    '    End If
    'Next

    If UCase(Right(rdata, Len(rdata) - 7)) = "TODOS" Then
        Call SendData(ToIndex, Userindex, 0, "||General: /ONLINE, /ONLINEMAP, /SALIR, /DESCANSAR, /MEDITAR, /RESUCITAR, /CURAR, /EST" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BUG, /MOVER, /DESC, /PASSWD, /CREDITOS, /POWA, /COMERCIAR, \, /MANTENIMIENTO" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||Clan: /FUNDARCLAN, /CMSG, /VOTO, /SALIRCLAN, /ONLINECLAN, ., -" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||Torneos: /TORNEO" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||Banco: /BALANCE, /BOVEDA, /DEPOSITAR, /RETIRAR" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||Mascotas: /ACOMPAÑAR, /QUIETO" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||Facciones: /ENLISTAR, /INFORMACION, /RECOMPENSA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Ayudante = False Then
            Call SendData(ToIndex, Userindex, 0, "||Usuarios: /CASAR, /DIVORCIARSE, /COMERCIAR, /REGALAR, /PARTY, /DEJARPARTY, /PMSG" & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, Userindex, 0, "||Usuarios: /CASAR, /DIVORCIARSE, /COMERCIAR, /REGALAR, /PARTY, /DEJARPARTY, /PMSG, /SHOW SOS, /SOS" & FONTTYPE_INFO)
        End If
        Call SendData(ToIndex, Userindex, 0, "||Creaturas o NPCs: /ENTRENAR, /COMERCIAR, /APOSTAR, /OTRACARA, /LOTERIA, /AVENTURA" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||GMs/Administradores: /GM, /GMSG, /URGENTE o /DENUNCIAR" & FONTTYPE_INFO)
        If AaP(Userindex) Then
            Call SendData(ToIndex, Userindex, 0, "||Para ver los comandos exclusivos de los Aprendizes de Administrador, /AYUDA APRENDIZ" & FONTTYPE_INFO)
        ElseIf EsAdmin(Userindex) Then
            Call SendData(ToIndex, Userindex, 0, "||Para ver los comandos exclusivos de los Administradores, /AYUDA ADMINISTRADOR" & FONTTYPE_INFO)
        End If
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub ' hasta aqui llega un usuario
        Call SendData(ToIndex, Userindex, 0, "||Para ver los comandos exclusivos de los Consejeros, /AYUDA CONSEJEROS" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Para ver los comandos exclusivos de los SemiDioses, /AYUDA SEMIDIOSES" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Para ver los comandos exclusivos de los Dioses, /AYUDA DIOSES" & FONTTYPE_INFO)
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "CONSEJEROS" Then
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then
            Call SendData(ToIndex, Userindex, 0, "||Los Consejeros son los encargados de ayudar aconsejando a los jugadores." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, Userindex, 0, "||General: /REM, /HORA, /TELEPLOC, /INVISIBLE, /INFOEST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Torneos: /ORGTORNEO, /FINTORNEO, /CUENTAREGRESIVA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Usuarios: /DONDE, /TELEP, /SHOW SOS, /CARCEL, /IRA, /CONSULTA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Creaturas o NPCs: /NENE" & FONTTYPE_INFO)
        End If
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "SEMIDIOSES" Then
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then
            Call SendData(ToIndex, Userindex, 0, "||Los SemiDioses son los encargados de mantener el orden y la disiplina de los jugadores." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, Userindex, 0, "||General: /ONLINEGM, /RMSG, /XMSG, /MSG, /N, /ROJO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Torneos: /GENTETORNEO, /MAPATORNEO, /YATORNEO, /NOTORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Usuarios: /INFO, /INV, /SKILLS, /REVIVIR, /PERDON, /ECHAR, /SILENCIAR, /FORCEMIDI, /FORCEWAV" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||/BAN, /UNBAN, /SUM, /IPNICK, /NICKIP, /VEN, /VS, /VSX, /CARXEL, /BOV, /BAL, /STAT" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Creaturas o NPCs: /SEGUIR, /CC, /RESETINV" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||GMs/ADMINITRADORES: /ONLINEGM" & FONTTYPE_INFO)
        End If
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "DIOSES" Then
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then
            Call SendData(ToIndex, Userindex, 0, "||Los Dioses son los dioses del juego, por lo tanto deven ser respetados y obedecidos." & FONTTYPE_INFO)
        ElseIf EsAdmin(Userindex) = True Or UserList(Userindex).flags.Privilegios >= 3 Then
            Call SendData(ToIndex, Userindex, 0, "||General: /HACERITEM (/CI), /CT, /CTT, /DT, /BLOQ, /LIMPIAR, /SMSG, /NAVE, /APAGAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||/DOBACKUP, /BORRAR SOS, /SHOW INT, /PASSDAY, /CLEANMAP(/CLEARMAP), /MASMANTENIMIENTO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||/BORRAR, /MASSDEST, /QUEST, /MAPAAGITE, /NOMAPAGITE, /BACK, /SEGURO, /HABILITAR, /PISO, /MASMANTENIMIENTO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Clan: /NOCLAN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Torneos: /RES, /MICROFONO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Facciones: /RAJA, /NOCASO, /NOREAL" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Usuarios: /BANIP, /UNBANIP, /CONDEN, /RAJAR, /MOD, /GRABAR, /MUSER, /CEGUERA, /LOG" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||/INVI, /SORRY, /KILLCHAR, /NIVELES, /BAS, /AYUDANTE, /NOEXISTE, /QUITAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Creaturas o NPCs: /DEST, /MATA, /MASSKILL, /ACC, /RACC, /AQUIAVENTURA, /NOAVENTURA, /NOCOUNTER, /COUNTER1, /COUNTER2, /COUNTER, /MASSKILLGROX, /ENCUESTA, /RENCUESTA, /BACKNPC, /PRETORIAN, /ACEPTCONSE, /ACEPTCONSECAOS, /KICKCONSE" & FONTTYPE_INFO)
        End If
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "GLOSARIO" Then
        Call SendData(ToIndex, Userindex, 0, "||Glosario:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||NPC = Personaje de inteligencia artificial, como creaturas y otros personajes." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||GM = GameMasters o administradores del juego." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||SKILLS = Habilidades de un personaje." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||SPAWN = Traer algo, por ejemplo SPAWN de Creaturas." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||RESPAWN = Nos trae automaticamente creaturas." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||NEWBIE = Novato, principante, aprendiz, etc." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||FACCIONES = Tambien podria llamarse bando, en el juego existen dos, Kaos o Armada." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||POWA = El mas poderoso, sabio, experimentado, etc." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||PT = Forma de llamar a los que se comportan peor que un novato." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||AGITE = Forma de llamar a un lugar en donde hay mucha gente con quien combatir." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||A CUCHI = Forma de llamar a un combate con cuchillos, espadas, etc, no con magias." & FONTTYPE_INFO)
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "GENERAL" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria GENERAL:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ONLINE, /ONLINEMAP, /SALIR, /DESCANSAR, /MEDITAR, /RESUCITAR, /CURAR, /EST" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BUG, /MOVER, /DESC, /PASSWD, /CREDITOS, /POWA, /COMERCIAR, \, ., -, /MANTENIMIENTO" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub ' hasta aqui llega un usuario
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Consejeros:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/REM, /HORA, /TELEPLOC, /INVISIBLE, /INFOEST" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo SemiDioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ONLINEGM, /RMSG, /XMSG, /MSG, /N, /ROJO" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/HACERITEM (/CI), /CT, /CTT, /DT, /BLOQ, /LIMPIAR, /SMSG, /NAVE, /APAGAR" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/DOBACKUP, /BORRAR SOS, /SHOW INT, /PASSDAY, /CLEANMAP(/CLEARMAP), /MASMANTENIMIENTO" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BORRAR, /MASSDEST, /QUEST, /MAPAAGITE, /NOMAPAGITE, /BACK, /SEGURO, /HABILITAR, /PISO" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "CLAN" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria CLAN:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/FUNDARCLAN, /CMSG, /VOTO, /SALIRCLAN, /ONLINECLAN" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/NOCLAN" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "TORNEOS" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria TORNEOS:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/TORNEO" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un usuario
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Consejeros:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ORGTORNEO, /FINTORNEO, /CUENTAREGRESIVA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo SemiDioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/GENTETORNEO, /MAPATORNEO, /YATORNEO, /NOTORNEO" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/RES, /MICROFONO" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "BANCO" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria BANCO:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BALANCE, /BOVEDA, /DEPOSITAR, /RETIRAR" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "MASCOTAS" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria MASCOTAS:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ACOMPAÑAR, /QUIETO" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "FACCIONES" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria FACCIONES:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ENLISTAR, /INFORMACION, /RECOMPENSA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/RAJA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/NOREAL, /NOCAOS" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "USUARIOS" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria USUARIOS:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/CASAR, /DIVORCIARSE, /COMERCIAR, /REGALAR, /PARTY, /DEJARPARTY, /PMSG" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un usuario
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Consejeros:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/DONDE, /TELEP, /SHOW SOS, /CARCEL, /IRA, /CONSULTA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo SemiDioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/INFO, /INV, /SKILLS, /REVIVIR, /PERDON, /ECHAR, /SILENCIAR, /FORCEMIDI, /FORCEWAV" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BAN, /UNBAN, /SUM, /IPNICK, /NICKIP, /VEN, /VS, /VSX, /CARXEL, /BOV, /BAL, /STAT" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/BANIP, /UNBANIP, /CONDEN, /RAJA, /MOD, /GRABAR, /MUSER, /CEGUERA" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/INVI, /SORRY, /KILLCHAR, /BAS, /NIVELES, /LOG, /AYUDANTE, /NOEXISTE, /QUITAR" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "CREATURAS" Or UCase(Right(rdata, Len(rdata) - 7)) = "NPCS" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria CREATURAS o NPCS:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ENTRENAR, /COMERCIAR, /APOSTAR, /OTRACARA, /LOTERIA, /AVENTURA" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un usuario
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Consejeros:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/NENE" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo SemiDioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/SEGUIR, /CC, /RESETINV, /RESET" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/DEST, /MATA, /MASSKILL, /ACC, /RACC, /AQUIAVENTURA, /NOAVENTURA, /NOCOUNTER, /COUNTER1, /COUNTER2, /COUNTER, /MASSKILLGROX, /ENCUESTA, /RENCUESTA, /BACKNPC, /PRETORIAN" & FONTTYPE_INFO)
        Exit Sub
    ElseIf UCase(Right(rdata, Len(rdata) - 7)) = "GMS" Or UCase(Right(rdata, Len(rdata) - 7)) = "ADMINISTRADORES" Then
        Call SendData(ToIndex, Userindex, 0, "||Categoria GMS/ADMINISTRADORES:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/GM, /GMSG, /URGENTE o /DENUNCIAR" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un consejero
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo SemiDioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/ONLINEGM" & FONTTYPE_INFO)
        If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
        Call SendData(ToIndex, Userindex, 0, "||Exclusivo Dioses:" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||/MASMANTENIMIENTO, /ACEPTCONSE, /ACEPTCONSECAOS, /KICKCONSE" & FONTTYPE_INFO)
        Exit Sub
    End If
    ' [/CATEGORIAS/]

    ' [COMANDOS]
    Select Case UCase(Right(rdata, Len(rdata) - 7))
        Case "/ONLINEMAP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ONLINEMAP" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra todos los nombres de los usuarios que esten en nuestro mismo mapa." & FONTTYPE_INFO)
        Case "/ONLINECLAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ONLINECLAN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra los nombres de los integrantes de nuestro clan que se encuentren online." & FONTTYPE_INFO)
        Case "/ONLINE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ONLINE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra todos los nombres de los usuarios que se encuentren jugando en el momento." & FONTTYPE_INFO)
        Case "/SALIR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SALIR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Sirve para salir correctamente del juego, dependiendo el servidor nos puede hacer esperar unos segundos antes de salir." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||No se puede salir mientras nos encontramos paralizados." & FONTTYPE_INFO)
        Case "/FUNDARCLAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /FUNDARCLAN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite fundar un clan. Necesitamos +90 de Skills en Liderasgo y ser mayor de nivel 20." & FONTTYPE_INFO)
        Case "/BALANCE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BALANCE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para utilizarlo, clickea en el banquero, y luego escribe este comando, para que el banquero nos diga cuanto dinero tenemos depositado." & FONTTYPE_INFO)
        Case "/QUIETO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /QUIETO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a nuestra mascota, ella se quedara quieta, y no nos seguira." & FONTTYPE_INFO)
        Case "/ACOMPAÑAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ACOMPAÑAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a nuestra mascota, ella comenzara a seguirnos a donde vallamos." & FONTTYPE_INFO)
        Case "/ENTRENAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ENTRENAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a un entrenador, nos mostrara una ventana en donde nos dara a elegir con que creatura deseamos entrenar." & FONTTYPE_INFO)
        Case "/DESCANSAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DESCANSAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para utilizarlo, debes situarte muy cerca de una fogata, al iniciar en modo de descanso, tu energia subira mucho mas rapido." & FONTTYPE_INFO)
        Case "/MEDITAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MEDITAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Cuando tenemos bajo el mana, de las magias, podemos inciar el modo de meditacion, para poder rellenar nuestro mana nuevamente." & FONTTYPE_INFO)
        Case "/RESUCITAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RESUCITAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a un sacerdote, el nos resucitara, si estamos vivos, no." & FONTTYPE_INFO)
        Case "/CURAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CURAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a un sacerdote, el nos curara la vida." & FONTTYPE_INFO)
        Case "/EST"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /EST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra nuestras estadisticas personales." & FONTTYPE_INFO)
        Case "/COMERCIAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /COMERCIAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Despues de haber clickeado un NPC o un usuario, pones este comando para iniciar el modo de comercio." & FONTTYPE_INFO)
        Case "/BOVEDA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BOVEDA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de clickear en un banquero, el nos dejara guardar items de nuestro inventario en su boveda." & FONTTYPE_INFO)
        Case "/ENLISTAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ENLISTAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Lo utilizaremos despues de clickear, al NPC que elijas como bando, recurda los bandos son Kaos y Armada." & FONTTYPE_INFO)
        Case "/INFORMACION"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INFORMACION" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Lo utilizaremos despues de clickear a tu jefe de bando. Y con esto nos dara mas informacion." & FONTTYPE_INFO)
        Case "/RECOMPENSA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RECOMPENSA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despus de haber clickeado a tu jefe de bando, no dira cuando reciviremos una recompensa o sino nos la dara en el momento." & FONTTYPE_INFO)
        Case "/CMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos sirve para comunicarnos con todos los integrantes de nuestro mismo clan." & FONTTYPE_INFO)
        Case "/GM"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /GM" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un pedido de ayuda al GM. Quien no necesariamente respondera rapido." & FONTTYPE_INFO)
        Case "/BUG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BUG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje sobre un error, que es archivado, y utilizado para corregir errores en el juego." & FONTTYPE_INFO)
        Case "/DESC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DESC <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite cambiar la descripcion de nuestro personaje." & FONTTYPE_INFO)
        Case "/VOTO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /VOTO <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vota a un integrante de tu clan, para lider. Solo si es dia de elecciones." & FONTTYPE_INFO)
        Case "/TORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /TORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al ponerlo una vez, nos inscribe a un torneo, si lo escribimos nuevamente nos borramos del torneo." & FONTTYPE_INFO)
        Case "/PASSWD"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PASSWD <password>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite cambiar el password de nuestro personaje." & FONTTYPE_INFO)
        Case "/RETIRAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RETIRAR <cantidad>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al uzarlo despues de haber clickeado un banquero, nos permitira retirar la cantidad de dinero que queramos, y tengamos depositada." & FONTTYPE_INFO)
        Case "/DEPOSITAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DEPOSITAR <cantidad>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado un banquero, nos permite depositar nuestro dinero en el banco." & FONTTYPE_INFO)
        Case "\"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: \<nombre> <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje privado al usuario que queramos." & FONTTYPE_INFO)
        Case "/GMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /GMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje directo a todos los GMs. Con lo que nos atenderan mucho mas rapido." & FONTTYPE_INFO)
        Case "-"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: -<mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos sirve para gritar un mensaje, en vez de hablarlo." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
        Case "/APOSTAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /APOSTAR <cantidad>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a un apostador, nos dejara apostar nuestro dinero a nuestra suerte." & FONTTYPE_INFO)
        Case "/SALIRCLAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SALIRCLAN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Con este comandos salimos de un clan, en el caso de que nosotros seamos el lider fundador y no hallan integrantes el clan es borrado y podras volver a fundar otro clan." & FONTTYPE_INFO)
        Case "/CREDITOS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CREDITOS" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Te informa sobre quienes ayudaron ha hacer el servidor." & FONTTYPE_INFO)
        Case "/CASAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CASAR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia una solicitud de matrimonio a el otro usuario, el cual deve de ser de distinto genero." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para aceptar tu solicitud la otra persona devera poner lo mismo pero con tu nombre." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para divorciarse, escribe /DIVORCIARSE" & FONTTYPE_INFO)
        Case "/DIVORCIARSE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DIVORCIARSE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos divorciamos con quien nos hallamos casado, sin necesidad de que la otra persona este jugando." & FONTTYPE_INFO)
        Case "/MOVER"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MOVER <lugar>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos teletransporta a una ciudad a cambio de dinero." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Los lugares y sus precios son: ULLA $" & str(MoverUlla) & ", NIX $" & str(MoverNix) & ", BANDER $" & str(MoverBander) & ", LINDOS $" & str(MoverLindos) & " y VERIL $" & str(MoverVeril) & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||No puedes teletransportante si estas paralizado." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Solo puedes teletransportarte si estas en un lugar seguro." & FONTTYPE_INFO)
        Case "/REGALAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /REGALAR <cantidad>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Regala una cantidad de dinero al ultimo jugador que hallamos clickeado." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos puede subir skilles en suerte." & FONTTYPE_INFO)
        Case "/URGENTE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /URGENTE <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje de urgencia al GM." & FONTTYPE_INFO)
        Case "/DENUNCIAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DENUNCIAR <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje de urgencia al GM." & FONTTYPE_INFO)
        Case "/OTRACARA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /OTRACARA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado sobre un NPC de reconstruccion facial nos cambiara a otra cara al azar por solo $" & str(ReconstructorFacial) & "." & FONTTYPE_INFO)
        Case "/POWA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /POWA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice quien es el jugador con más nivel en todo el servidor, cual es el más PK y quien es el que más online." & FONTTYPE_INFO)
        Case "/PARTY"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PARTY" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de haber clickeado a un jugador, solicitamos o aceptamos una peticion de party, que consiste en dividir las experiencias." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para quitarlo, necesitamos poner /DEJARPARTY." & FONTTYPE_INFO)
        Case "/DEJARPARTY"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DEJARPARTY" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo saldremos de la party en donde estemos, ya sea como lider o participante." & FONTTYPE_INFO)
        Case "."
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: <parte1>.<parte2>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos sirve para representar un espacio en el nombre de un personaje." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Se puede representar infinitas veces." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo: \el.chavo Hola." & FONTTYPE_INFO)
        Case "/LOTERIA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /LOTERIA <num1> <num2>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de clickear sobre un NPC de loteria, podemos participar en ella," & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||por el pozo acumulado, que supera el millones de monedas de oro." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Solo tenemos que elegir 2 numeros de 2 cifras y porbar nuestra destreza por solo " & str(BoletoDeLoteria) & " monedas de oro." & FONTTYPE_INFO)    ' GS-Server Commands
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo: /LOTERIA 93 00" & FONTTYPE_INFO)    ' GS-Server Commands
        Case "/AVENTURA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /AVENTURA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de clickear sobre un NPC aventurero, nos llevara a una aventura con un costo de " & BoletoAventura & " monedas de oro." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Saldras de ella cuando pase el tiempo especificado, o si mueres." & FONTTYPE_INFO)
        ' Version v0.12t
        Case "/PMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia el mensaje a todos los integrantes de la Party." & FONTTYPE_INFO)
        ' Version v0.12t8
        Case "/SHOW SOS"
            If UserList(Userindex).flags.Ayudante = True Then
                Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SHOW SOS" & FONTTYPE_INFO)
                Call SendData(ToIndex, Userindex, 0, "||Nos muestra los usuarios que estan pidiendo ayuda y el mapa donde se encuentran." & FONTTYPE_INFO)
            End If
        Case "/SOS"
            If UserList(Userindex).flags.Ayudante = True Then
                Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SOS <numero>" & FONTTYPE_INFO)
                Call SendData(ToIndex, Userindex, 0, "||Quitamos un pedido de ayuda." & FONTTYPE_INFO)
            End If
        ' v0.12b1
        Case "/MANTENIMIENTO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MANTENIMIENTO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dise cuanto tiempo falta para el proximo Mantenimiento." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    End Select
    If UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un usuario
    ' Comando permitido?
    If ComandoPermitido(Userindex, UCase(Right(rdata, Len(rdata) - 7))) = False Then
        If EsAdmin(Userindex) = True Then
            Call SendData(ToIndex, Userindex, 0, "||Sus permisos no son validos para ver la ayuda de este comando." & FONTTYPE_INFX)
            Exit Sub
        End If
    End If
    Select Case UCase(Right(rdata, Len(rdata) - 7))
        Case "/REM"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /REM <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos deja dejar mensajes para los administradores." & FONTTYPE_INFO)
        Case "/HORA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /HORA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice la hora, segun el servidor y tambien se la dice a todos los jugadores." & FONTTYPE_INFO)
        Case "/DONDE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DONDE <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice en que mapa, y posicion se encuentra un usuario." & FONTTYPE_INFO)
        Case "/NENE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NENE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice cuantos NPC hay en el mapa." & FONTTYPE_INFO)
        Case "/TELEPLOC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /TELEPLOC" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos teletransportamos al ultimo lugar donde hallamos hecho click." & FONTTYPE_INFO)
        Case "/TELEP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /TELEP <nombre> <mapa> <pos-x> <pos-y>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Teletranportamos al usuario que querramos al un mapa, en la posicion que querramos." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para teletranportarnos nosotros, ponemos YO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo: /TELEP yo 1 50 50" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Este ejemplo, nos teletransorta a Ulla(1) en la posicion 50 50." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||El maximo de posicion es 90 90." & FONTTYPE_INFO)
        Case "/SHOW SOS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SHOW SOS" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra la lista de peticiones de auxilio de los usuarios que hallan hecho /GM" & FONTTYPE_INFO)
        Case "/ORGTORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ORGTORNEO <tipo>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos deja realizar automaticamente presentaciones de distintos tipos de torneo." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Los tipos son: 1-Duelo 2-Guerra 3-Eliminatorias" & FONTTYPE_INFO)
        Case "/FINTORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /FINTORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra la lista de peticiones de torneo." & FONTTYPE_INFO)
        Case "/CUENTAREGRESIVA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CUENTAREGRESIVA <numero>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Realiza una cuenta regresiva. El maximo es 99." & FONTTYPE_INFO)
        Case "/IRA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /IRA <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos teletransporta a donde se cuentra el usuario." & FONTTYPE_INFO)
        Case "/INVISIBLE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INVISIBLE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos hace invisibilidad de administracion. Para espiar a los usuarios." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
        Case "/INFOEST"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INFOEST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra informacion de estado de los modos del servidor." & FONTTYPE_INFO)
        Case "/CONSULTA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CONSULTA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Iniciamos modo consulta, y todo usuario y creatura en nuestra pantalla no podra atacar." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Solo puede utilizar este modo un administrador a la vez." & FONTTYPE_INFO)
        ' Version t0.12
        Case "/RELOAD"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RELOAD" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Recarga los permisos de administracion." & FONTTYPE_INFO)
        Case "/CARCEL"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CARCEL <tiempo> <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia a la carcel a un usuario por un tiempo determinado en minutos." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||El maximo de tiempo es 60 minutos." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
    End Select
    If UserList(Userindex).flags.Privilegios < 2 And EsAdmin(Userindex) = False Then Exit Sub  ' hasta aqui llega un consejero
    Select Case UCase(Right(rdata, Len(rdata) - 7))
        Case "/INFO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INFO <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra las estadisticas del usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/INV"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INV <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra el inventario de un usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/SKILLS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SKILLS <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra los skilles de un usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/REVIVIR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /REVIVIR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite revivir a un usuario." & FONTTYPE_INFO)
        Case "/ONLINEGM"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ONLINEGM" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra los nombres de los GMs que se encuentren online en el momento." & FONTTYPE_INFO)
        Case "/PERDON"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PERDON <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vuelve ciudadano al cualquier usuario." & FONTTYPE_INFO)
        Case "/ECHAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ECHAR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un usuario momentaneamente." & FONTTYPE_INFO)
        Case "/BAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BAN <razon>@<nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un usuario y lo deja inavilitado para volver a usar el personaje." & FONTTYPE_INFO)
        Case "/UNBAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /UNBAN <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vuelve a permitirle ingresar con el personaje a un usuario previamente Baneado." & FONTTYPE_INFO)
        Case "/SEGUIR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SEGUIR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizarlo despues de clickear a una creatura, esta nos seguira." & FONTTYPE_INFO)
        Case "/SUM"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SUM <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Trae hacia nosotros al usuario que deseemos." & FONTTYPE_INFO)
        Case "/CC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CC" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra una lista de creaturas que podemos spawnear." & FONTTYPE_INFO)
        Case "/RESETINV"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RESETINV" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vacia el inventario del NPC que hallamos clickeado anteriormente." & FONTTYPE_INFO)
        Case "/RMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje ha todos los usuarios." & FONTTYPE_INFO)
        Case "/IPNICK"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /IPNICK <ip>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice cual de los usuarios tiene ese IP." & FONTTYPE_INFO)
        Case "/NICKIP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NICKIP <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice cual es el IP del usuario." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
        Case "/MSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje exclusivamente para otros GMs, osea no lo ve todo el mundo." & FONTTYPE_INFO)
        Case "/VEN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /VEN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Trae hacia nosotros el ultimo usuario que hallamos clickeado." & FONTTYPE_INFO)
        Case "/N"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /N" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Debug de lo que hallamos que hallamos clickeado, usuarios, objetos y npc." & FONTTYPE_INFO)
        Case "/VS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /VS <num-ind>,<num-ind>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Presenta a los dos usuarios, en un Versus." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para conseguir el numero de index, hacer click en el usuario y poner /N" & FONTTYPE_INFO)
        Case "/VSX"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /VSX <num-ind>,<num-ind>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Trea a los dos usuarios, uno a nuestra izquerda y el otro a la derecha, y despues de esto los presenta en un Versus." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Para conseguir el numero de index, hacer click en el usuario y poner /N" & FONTTYPE_INFO)
        Case "/CARXEL"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CARXEL <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia a la carcel a un usuario, sin necesidad que este se encuentre jugando." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||El maximo tiempo de la pena es de 60 minutos." & FONTTYPE_INFO)
        Case "/GENTETORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /GENTETORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Transporta a toda los los usuarios que esten inscriptos a un torneo, los deja al azar hasta 3 pasos bajo nosotros." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Si uno o varios usuarios ya se encuentran en el mapa del torneo, no los teletransporta, solo gente nueva." & FONTTYPE_INFO)
        Case "/YATORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /YATORNEO <AUTO>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Iniciamos el modo torneo, automaticamente indicamos que mapa de torneos es el actual." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Si colocamos /YATORNEO AUTO, activara el AutoComentarista, para desactivarlo, desactiva el Torneo." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Previamente un dios tendria que haber configurado las restricciones del torneo." & FONTTYPE_INFO)
        Case "/MAPATORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MAPATORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Indicamos que el mapa actual es un mapa de torneo." & FONTTYPE_INFO)
        Case "/NOTORNEO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOTORNEO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Damos por terminado el torneo, las restriciones se desactivan y el mapa indicado para torneo es borrado." & FONTTYPE_INFO)
        Case "/ROJO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ROJO <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Mostramos un mensaje en rojo en la consola, para todos los usuarios." & FONTTYPE_INFO)
        Case "/RESET"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RESET" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Resetea un NPC cualquiera que hallamos clickeado anteriormente, sirve para los npc comerciantes." & FONTTYPE_INFO)
        Case "/XMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /XMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envía un mensaje ha todos los usuarios, sin nombre de personaje. En color azul marino." & FONTTYPE_INFO)
        ' 0.12t8
        Case "/BOV"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BOV <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra la boveda de un usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/BAL"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BAL <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra el balance de un usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/STAT"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /STAT <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra las mini-estadisticas de un usuario. (Online o Offline)" & FONTTYPE_INFO)
        ' 0.12b1
        Case "/SILENCIAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SILENCIAR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite silenciar cualquier usuario, este no podra mandar /GMSG, /DENUNCIAR o enviar privados a cualquier GM's." & FONTTYPE_INFO)
        Case "/FORCEMIDI"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /FORCEMIDI <midi>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite hacer sonar un Midi(Musica) en todo el Mundo. (del 1 a 21 y desde 50 a 55)" & FONTTYPE_INFO)
        Case "/FORCEWAV"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /FORCEWAV <wav>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos permite escuchar un Wav(Sonido) en todo el Mundo. (del 2 a 108/Fuego/Click/Harp3/Cupdice/Click2/etc...)" & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
    End Select
    If UserList(Userindex).flags.Privilegios < 3 Or AaP(Userindex) = True Then Exit Sub  ' hasta aqui llega un semi dios
    Select Case UCase(Right(rdata, Len(rdata) - 7))
        Case "/HACERITEM"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /HACERITEM <cantidad>@<num-obj>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Permite crear objetos en cantidades." & FONTTYPE_INFO)
        Case "/CI"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CI <cantidad>@<num-obj>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Permite crear objetos en cantidades." & FONTTYPE_INFO)
        Case "/BANIP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BANIP <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un usuario y bloque su IP, impidiendo volver a conectarse al servidor con otros personajes." & FONTTYPE_INFO)
        Case "/UNBANIP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /UNBANIP <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vuelve a permitirle entrar al usuario baneado por IP." & FONTTYPE_INFO)
        Case "/CT"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CT <map> <pos-x> <pos-y>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crear un teletransporte al mapa y la posicion que querramos que lleve. Tambien donde mapa, se puede usar Ulla, Lindos, Bander o Nix, pero sin posiciones." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 1: /CT 1 50 50" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 2: /CT Lindos" & FONTTYPE_INFO)
        Case "/CTT"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CTT <map> <pos-x> <pos-y>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crear un teletransporte transparente al mapa y la posicion que querramos que lleve. Tambien donde mapa, se puede usar Ulla, Lindos, Bander o Nix, pero sin posiciones." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 1: /CTT 1 50 50" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 2: /CTT Lindos" & FONTTYPE_INFO)
        Case "/DT"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DT" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Destruye el ultimo portal al que hallamos clickeado." & FONTTYPE_INFO)
        Case "/DEST"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DEST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Destruye el objeto sobre el que estemos parados." & FONTTYPE_INFO)
        Case "/BLOQ"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BLOQ" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Hace un obstaculo en donde estemos parados. Bloquea y desbloquea donde estemos parados." & FONTTYPE_INFO)
        Case "/MATA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MATA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra el ultimo NPC que hallamos clickeado." & FONTTYPE_INFO)
        Case "/MASSKILL"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MASSKILL" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra todos los NPC de el mapa." & FONTTYPE_INFO)
        Case "/LIMPIAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /LIMPIAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Quita todos los NPC del mundo." & FONTTYPE_INFO)
        Case "/SMSG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SMSG <mensaje>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Envia un mensaje de sistema, en una ventana aparte." & FONTTYPE_INFO)
        Case "/ACC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ACC <npc-num>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crea un NPC." & FONTTYPE_INFO)
        Case "/RACC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RACC <npc-num>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crea un NPC con ReSpawn." & FONTTYPE_INFO)
        Case "/NAVE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NAVE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Depuracion del mar." & FONTTYPE_INFO)
        Case "/APAGAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /APAGAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Apaga el servidor!!!" & FONTTYPE_INFO)
        Case "/CONDEN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CONDEN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Vuelve criminal a un usuario que hallamos clickeado." & FONTTYPE_INFO)
        Case "/RAJAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RAJAR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Le quita todas las facciones a un usuario." & FONTTYPE_INFO)
        Case "/MOD"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MOD <nombre> <tipo> <valor>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Modifica a un usuario. Los tipos son ORO(puede usarse MAX), EXP(puede usarse MAX), BODY(edita el cuerpo), HEAD(edita la cabeza), CRI(cantidad de criminales matados), CIU(cantidad de ciudadanos matados), LEVEL(nivel), GEN(genero, (2)MUJER o (1)HOMBRE), HP(vida), MP(mana), ST(Stamina), HIT(fuerza, DEF(defensa), EMAIL, NICK, SKILL, DADOS." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 1: /MOD GS ORO 99999999" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 2: /MOD TRABA GEN 2" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo 2: /MOD ALGUIEN ORO MAX" & FONTTYPE_INFO)
        Case "/DOBACKUP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /DOBACKUP" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Guarda el estado las principales ciudades y los NPC." & FONTTYPE_INFO)
        Case "/GRABAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /GRABAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Graba los personajes." & FONTTYPE_INFO)
        Case "/BORRAR SOS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BORRAR SOS" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra la lista de peticiones de ayuda." & FONTTYPE_INFO)
        Case "/SHOW INT"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SHOW INT" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||???? Hace algo en el servidor, no es nada util." & FONTTYPE_INFO)
        Case "/LLUVIA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /LLUVIA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Comienza o detiene la lluvia." & FONTTYPE_INFO)
        Case "/PASSDAY"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PASSDAY" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Hace pasar un dia. Nos puede servir sabiendo que los clanes votan a su lider despues de unos dias." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
        Case "/CLEANMAP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CLEANMAP o /CLEANMAP <mapa>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Limpiamos el mapa donde estemos de todos los items no pertenecientes al mapa, por ejemplo oro." & FONTTYPE_INFO)
        Case "/CLEARMAP"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CLEARMAP" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Limpiamos el mapa donde estemos de todos los items no pertenecientes al mapa, por ejemplo oro." & FONTTYPE_INFO)
        Case "/MUSER"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MUSER <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Matamos repentinamente a un usuario. Tambien funciona con la palabra YO para matarnos a nosotros mismos." & FONTTYPE_INFO)
        Case "/CEGUERA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /CEGUERA <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Activa o desactiva la ceguera sobre un usuario. Por tiempo indeterminado." & FONTTYPE_INFO)
        Case "/INVI"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /INVI <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Activa o desactiva invisibilidad sobre un usuario. Por tiempo indeterminado." & FONTTYPE_INFO)
        Case "/NOCLAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOCLAN <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra el clan a un usuario, asi como informacion hacerca de que ya ha fundado uno." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Permite al usuario voler a fundar otro clan." & FONTTYPE_INFO)
        Case "/RES"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RES <tipo> <pots> <mascotas> <no ko> <no se caen items>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Configura las restriciones de un torneo." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Los tipos son 1-Vale Todo, 2-No ceguera, invi o estupidez, 3-Idem pero sin para o inmo, 4-Sin magias a Cuchi" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||En donde pots, 1-si se permiten, 0-no" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||En donde no KO, 1-si es que no esta permitido, 0-si esta permitido" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||KO significa los golpes o hehcizos que bajan de una." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||En donde mascotas va el maximo de mascotas que se permiten, 0=ninguna" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Y si en donde no se caen items es 1, los usuarios al morir en el torneo no tiran los items." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Las restricciones son guardadas, asi que no requieren volver a configurarse muy seguido." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Ejemplo: /RES 1 1 3 0 0" & FONTTYPE_INFO)
        Case "/SORRY"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SORRY <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Quita el tiempo el prision del usuario y lo teletransporta a Ulla." & FONTTYPE_INFO)
        Case "/BORRAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BORRAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra el ultimo objeto que hallamos clickeado, tambien puertas y otros." & FONTTYPE_INFO)
        Case "/MASSDEST"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MASSDEST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra todo el oro en el mapa." & FONTTYPE_INFO)
        Case "/KILLCHAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /KILLCHAR <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra el archivo del usuario, con lo que es peor que banear el personaje." & FONTTYPE_INFO)
        Case "/BAS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BAS <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Marca a un usuario, newbie, como personaje basura, y en cuanto desconecte es automaticamente eliminado." & FONTTYPE_INFO)
        Case "/QUEST"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /QUEST" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Activa y desactiva el modo Quest." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Los administradores, al ser clickeados(invisibles o no) por los usuarios informa por la consola, quien descubrio a quien." & FONTTYPE_INFO)
        Case "/MICROFONO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MICROFONO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Solo se puede utilizar en un torneo, y mientras se encuentre activado." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Hace que todo lo que digan los GMs salga en la consola para que lo lean todos." & FONTTYPE_INFO)
        Case "/MAPAAGITE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MAPAAGITE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Define un mapa, como mapa de agite, en el cual los usuarios tendran varias ventajas para que les guste el lugar." & FONTTYPE_INFO)
        Case "/NOMAPAAGITE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOMAPAAGITE" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Borra el mapa definido como mapa de agite." & FONTTYPE_INFO)
        Case "/NIVELES"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NIVELES" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos muestra el nombre y el nivel de todos los usuarios que se encuentren en el momento." & FONTTYPE_INFO)
        Case "/BACK"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BACK" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Guarda un backup del mapa en donde estemos. (Sin NPC, para guardar el BK de NPC's del mundo escribe /BACKNPC)" & FONTTYPE_INFO)
        Case "/LOG"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /LOG <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Guarda como ultima posicion del usuario la posicion en donde estemos." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos sirve para deslogear usuarios." & FONTTYPE_INFO)
        Case "/SEGURO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /SEGURO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Activa y desactiva, si el mapa es zona segura o no. No es permanente a menos que haga un /BACK." & FONTTYPE_INFO)
        Case "/AQUIAVENTURA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /AVENTURA <tiempo>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Determina el lugar donde estamos parados como el comienzo de la aventura, ademas de determinar el tiempo de tal, de 1 a 30 minutos." & FONTTYPE_INFO)
        Case "/NOAVENTURA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOAVENTURA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Deshabilita el mapa de la aventura." & FONTTYPE_INFO)
        Case "/COUNTER1"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /COUNTER1" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Determina el lugar donde estamos parados como el comienzo, del bando de los Criminal." & FONTTYPE_INFO)
        Case "/COUNTER2"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /COUNTER2" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Determina el lugar donde estamos parados como el comienzo, del bando de los Ciudadanos." & FONTTYPE_INFO)
        Case "/COUNTER"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /COUNTER <valor ingreso> <valor muerte>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Determina el mapa donde estamos parados, como mapa del Modo Counter." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Y configura el minimo de oro que necesitara el NPC para dejarnos participar, y cuanto oro reciviremos al matar a un contrincante." & FONTTYPE_INFO)
        Case "/NOCOUNTER"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOCOUNTER" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Deshabilita el mapa del modo Counter." & FONTTYPE_INFO)
        ' Version t0.12
        Case "/MASSKILLGROX"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MASSKILLGROX" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Mata todos los NPC's del mapa." & FONTTYPE_INFO)
        Case "/BACKNPC"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /BACKNPC" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crea un BackUp de NPC de todo el mundo." & FONTTYPE_INFO)
        Case "/ENCUESTA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ENCUESTA <encuesta>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crea una encuesta, si no colocamos encuesta Termina una encuesta iniciada." & FONTTYPE_INFO)
        Case "/RENCUESTA"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /RENCUESTA" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Repite la consigna de la encuesta." & FONTTYPE_INFO)
        Case "/PRETORIAN"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PRETORIAN" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Crea el Clan Pretoriano en el Mapa especificado como Pretoriano." & FONTTYPE_INFO)
        ' v012t8
        Case "/AYUDANTE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /AYUDANTE <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Le damos/quitamos el cargo de Ayudante a cualquier usuario. (Online o Offline)" & FONTTYPE_INFO)
        Case "/HABILITAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /HABILITAR" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Al utilizar este comando, si el servidor esta Reservado para Administradores, dejara de estarlo." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||En el caso que el servidor este para todo publico, expulsara a todos los usuarios sin privilegios y el servidor pasara a estar Reservado para Administradores." & FONTTYPE_INFO)
        ' 0.12b1
        Case "/NOEXISTE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOEXISTE <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Hace que el usuario deje de existir/exista en el juego, no tendra nombre, no se indicara los privilegios, no es clickeable y no aparecera entre los Online." & FONTTYPE_INFO)
        Case "/MASMANTENIMIENTO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /MASMANTENIMIENTO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Agrega una hora más al tiempo restante de Mantenimiento, cuando el Mantenimiento es menor que una hora." & FONTTYPE_INFO)
        Case "/ACEPTCONSE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ACEPTCONSE <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Aceptar al usuario en el Consejo Real." & FONTTYPE_INFO)
        Case "/ACEPTCONSECAOS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /ACEPTCONSECAOS <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Aceptar al usuario en el Consejo de las Sombras." & FONTTYPE_INFO)
        Case "/KICKCONSE"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /KICKCONSE <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un miembro del Consejo." & FONTTYPE_INFO)
        Case "/PISO"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /PISO" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Nos dice que objetos hay en el suelo de Todo el Mapa y su Cantidad." & FONTTYPE_INFO)
        Case "/NOREAL"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOREAL <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un miembro faccionario de la Armada y impide su reenlistacion." & FONTTYPE_INFO)
        Case "/NOCAOS"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /NOCAOS <nombre>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Expulsa a un miembro faccionario de las Sombras y impide su reenlistacion." & FONTTYPE_INFO)
        ' 0.12b3
        Case "/QUITAR"
            Call SendData(ToIndex, Userindex, 0, "||Syntaxis: /QUITAR <nombre>@<numobj>" & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||Quita el objeto indicado, del inventario y banco de cualquier usuario online." & FONTTYPE_INFO)
    ' GS-Server Commands
    ' GS-Server Commands
    ' GS-Server Commands
    End Select

End Sub
