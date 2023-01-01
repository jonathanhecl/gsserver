Attribute VB_Name = "GS_TCP_Logged"
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


' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos
' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos
' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos

Function TCP_Movimientos(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    ' "M" y "FPS"
    TCP_Movimientos = True
    ' [GS] NPC en rango
    If UserList(UserIndex).flags.SuNPC > 0 Then
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.SuNPC).Pos) > 8 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(UserList(UserIndex).flags.SuNPC).Name & " esta fuera de tu rango de vision." & FONTTYPE_INFO)
            UserList(UserIndex).flags.SuNPC = 0
                ' No es suyo, porque perdio la pertencia
        End If
    End If
    ' [/GS]
    Select Case UCase$(Left$(rdata, 1))
        Case "M" 'Moverse
        ' [NEW] Anti SH!!!!!! :O
            ' [GS] Repara el bug de la meditacion
                If UserList(UserIndex).flags.Meditando = True Then
                    ' Esta meditanto, se lo quito para no hacer lag
                    UserList(UserIndex).flags.Meditando = False
                    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
                    UserList(UserIndex).Char.FX = 0
                    UserList(UserIndex).Char.loops = 0
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
                    Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
                End If
            ' [/GS]
            ' [GS] Esta tirando una explocion magica??
                If UserList(UserIndex).flags.TiraExp = True Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.TiraExp = False
                End If
            ' [/GS]
            ' [GS] Esta en aventura?
                If UserList(UserIndex).flags.AV_Esta = True And MapaAventura <> UserList(UserIndex).Pos.Map Then
                    ' si esta en aventura y no esta en el mapa, termina hay la aventura
                    UserList(UserIndex).flags.AV_Tiempo = 0
                    UserList(UserIndex).flags.AV_Esta = False
                    'Call WarpUserChar(tIndex, mapa, X, Y, True)
                    ' Devolver
                    Call SendData(ToIndex, tIndex, 0, "||Tú aventura ha terminado." & "~255~255~0~1~0")
                End If
            ' [/GS]
            ' [GS] Esta en modo Counter?
                If UserList(UserIndex).flags.CS_Esta = True And MapaCounter <> UserList(UserIndex).Pos.Map Then
                    UserList(UserIndex).flags.CS_Esta = False
                End If
            ' [/GS]
            ' [GS] Bug del paralizado muerto
                If UserList(UserIndex).flags.Muerto = 1 Then
                    If UserList(UserIndex).flags.Paralizado = 1 Then
                        UserList(UserIndex).flags.Paralizado = 0
                        Call SendData(ToIndex, UserIndex, 0, "PARADOK")
                    End If
                    If UserList(UserIndex).flags.Meditando = True Then
                        UserList(UserIndex).flags.Meditando = False
                        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
                    End If
                End If
            ' [/GS]
            ' [GS] GM sin anti-sh
                If (UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex)) Or AntiSpeedHack = False Then GoTo PuedeSH:
            ' [/GS]
                Dim dummy As Long
                Dim TempTick As Long
                If UserList(UserIndex).flags.TimesWalk >= 30 Then
                    TempTick = GetTickCount And &H7FFFFFFF
                    dummy = (TempTick - UserList(UserIndex).flags.StartWalk)
                    If dummy < 6050 Then
                        If TempTick - UserList(UserIndex).flags.CountSH > 90000 Then
                            UserList(UserIndex).flags.CountSH = 0
                        End If
                        If Not UserList(UserIndex).flags.CountSH = 0 Then
                            dummy = 126000 / dummy
                            Call LogUsoSh(UserList(UserIndex).Name)
                            Call SendData(ToAll, 0, 0, "||<Anti-Chit> " & UserList(UserIndex).Name & " ha sido expulsado por posible uso de SpeedHack." & FONTTYPE_WARNING)
                            Call CloseSocket(UserIndex)
                            Exit Function
                        Else
                            UserList(UserIndex).flags.CountSH = TempTick
                        End If
                    End If
                    UserList(UserIndex).flags.StartWalk = TempTick
                    UserList(UserIndex).flags.TimesWalk = 0
                End If
                
                UserList(UserIndex).flags.TimesWalk = UserList(UserIndex).flags.TimesWalk + 1
                
                rdata = Right$(rdata, Len(rdata) - 1)
                
                If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando _
                   And UserList(UserIndex).flags.Paralizado = 0 Then
                      Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Descansar Then
                  UserList(UserIndex).flags.Descansar = False
                  Call SendData(ToIndex, UserIndex, 0, "DOK")
                  Call SendData(ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                  Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Meditando Then
                  UserList(UserIndex).flags.Meditando = False
                  Call SendData(ToIndex, UserIndex, 0, "MEDOK")
                  Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
                  UserList(UserIndex).Char.FX = 0
                  UserList(UserIndex).Char.loops = 0
                  Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
                  Call MoveUserChar(UserIndex, val(rdata))
                Else
                    '[CDT 17-02-2004]
                  If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 1
                  End If
                  '[/CDT]
                  UserList(UserIndex).flags.CountSH = 0
                End If
                
                If UserList(UserIndex).flags.Oculto = 1 Then
                    
                    If UCase$(UserList(UserIndex).clase) <> CLASS_LADRON Then
                        Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                        UserList(UserIndex).flags.Oculto = 0
                        UserList(UserIndex).flags.Invisible = 0
                        Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                    End If
                    
                End If
        
                Exit Function
        ' [/NEW]
        
        ' [OLD]
PuedeSH:
            rdata = Right$(rdata, Len(rdata) - 1)
           
            If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando _
               And UserList(UserIndex).flags.Paralizado = 0 Then
                  Call MoveUserChar(UserIndex, val(rdata))
            ElseIf UserList(UserIndex).flags.Descansar Then
              UserList(UserIndex).flags.Descansar = False
              Call SendData(ToIndex, UserIndex, 0, "DOK")
              Call SendData(ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
              Call MoveUserChar(UserIndex, val(rdata))
            ElseIf UserList(UserIndex).flags.Meditando Then
              UserList(UserIndex).flags.Meditando = False
              Call SendData(ToIndex, UserIndex, 0, "MEDOK")
              Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
              UserList(UserIndex).Char.FX = 0
              UserList(UserIndex).Char.loops = 0
              Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
              Call MoveUserChar(UserIndex, val(rdata))
            Else
              Call SendData(ToIndex, UserIndex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
            End If
           ' UserList(UserIndex).Counters.AntiSH = UserList(UserIndex).Counters.AntiSH + 1
            If UserList(UserIndex).flags.Oculto = 1 Then
           
                If UCase$(UserList(UserIndex).clase) <> CLASS_LADRON Then
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.Oculto = 0
                    UserList(UserIndex).flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                End If
           
            End If
           
            Exit Function
        ' [/OLD]
    End Select
    
    ' [NEW]
    If UCase$(Left$(rdata, 3)) = "FPS" Then
        rdata = Right$(rdata, Len(rdata) - 3)
        If Not haciendoBK Then UserList(UserIndex).CheatCont = UserList(UserIndex).CheatCont + 1
        If val(rdata) > 21 Then
            ' [GS]
            If (UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex)) Then
                Call LogCOSAS(UserList(UserIndex).Name, "Uso de cheats FPS: " & rdata, (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
                Exit Function
            End If
            ' [/GS]
            Call LogBan(UserIndex, UserIndex, "Uso de cheats FPS: " & rdata)
            Call SendData(ToAll, 0, 0, "||<Anti-Chit> Expulso a " & UserList(UserIndex).Name & "." & FONTTYPE_FIGHT)
            Call SendData(ToAll, 0, 0, "||<Anti-Chit> Baneo a " & UserList(UserIndex).Name & "." & FONTTYPE_FIGHT)
            UserList(UserIndex).flags.ban = 1
            Call CloseSocket(UserIndex)
        End If
        Exit Function
    End If
    ' [/NEW]
    TCP_Movimientos = False
End Function

' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos
' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos
' TCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_MovimientosTCP_Movimientos

' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones
' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones
' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones

Function TCP_Acciones(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    TCP_Acciones = True
    ' TI, LH, UK, RC, ATRI, FAMA, FEST, ESKI, WLC
    Select Case UCase$(Left$(rdata, 2))
        Case "TI" 'Tirar item
                If UserList(UserIndex).flags.Muerto = 1 Or _
                   (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)) Then Exit Function
                   '[Consejeros]
                   
                ' [GS] Counter mode?
                If UserList(UserIndex).flags.CS_Esta = True And UserList(UserIndex).Pos.Map = MapaCounter Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes tirar objetos en este mapa!" & "~255~255~0~1~0")
                    Exit Function
                End If
                ' [/GS]
                
                ' [GS] Es dios y navega si:P
                If UserList(UserIndex).flags.Navegando = 1 And _
                    (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) Then Exit Function
                ' [/GS]
                
                rdata = Right$(rdata, Len(rdata) - 2)
                Arg1 = ReadField(1, rdata, 44)
                Arg2 = ReadField(2, rdata, 44)
                If val(Arg1) = FLAGORO And val(Arg2) >= 5 Then
                    Call TirarOro(val(Arg2), UserIndex)
                    Call SendUserStatsBox(UserIndex)
                    UserList(UserIndex).flags.BugLageador = 0
                ElseIf val(Arg1) = FLAGORO Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes tirar tan poco oro, podrias producir lag." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.BugLageador = UserList(UserIndex).flags.BugLageador + 1
                    Exit Function
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Function
                        End If
                        Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                    Else
                        Exit Function
                    End If
                End If
                Exit Function
        Case "LH" ' Lanzar hechizo
            UserList(UserIndex).flags.BugLageador = 0
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Function
            End If
            ' [GS] Al fin :D
            If UserList(UserIndex).flags.Descansar Then Exit Function
            ' [/GS]
            rdata = Right$(rdata, Len(rdata) - 2)
            'UserList(Userindex).flags.PuedeLanzarSpell = 1
            UserList(UserIndex).flags.Hechizo = val(rdata)
            'Call SendData(ToIndex, Userindex, 0, "||Lanzar " & rdata & FONTTYPE_GS)
            Exit Function
        Case "LC" 'Click izquierdo
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Function
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Function
        Case "RC" 'Click derecho
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Function
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Function
        Case "UK"
            If UserList(UserIndex).flags.Muerto = 1 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 13 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 13
                End If
                Exit Function
            End If
            'Call SendData(ToIndex, Userindex, 0, "||Skill " & Magia & FONTTYPE_GS)
            rdata = Right$(rdata, Len(rdata) - 2)
            UserList(UserIndex).flags.QuiereLanzarSpell = 0
            Select Case val(rdata)
                Case Robar
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Robar)
                Case Magia
                    If UserList(UserIndex).flags.PuedeLanzarSpell = 0 Then
                        UserList(UserIndex).flags.QuiereLanzarSpell = 1
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "T01" & Magia)
                    End If
                Case Domar
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Domar)
                Case Ocultarse
                    
                    If UserList(UserIndex).flags.Navegando = 1 Then
                            If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                                Call SendData(ToIndex, UserIndex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                                UserList(UserIndex).flags.UltimoMensaje = 3
                            End If
                          Exit Function
                    End If
                    
                    If UserList(UserIndex).flags.Oculto = 1 Then
                              If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                                Call SendData(ToIndex, UserIndex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                                UserList(UserIndex).flags.UltimoMensaje = 2
                              End If
                          Exit Function
                    End If
                    
                    Call DoOcultarse(UserIndex)
            End Select
            Exit Function
    End Select
    If UCase$(Left$(rdata, 3)) = "WLC" Then 'Click izquierdo en modo trabajo
            rdata = Right$(rdata, Len(rdata) - 3)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            Arg3 = ReadField(3, rdata, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Function
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Function
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(UserIndex).flags.Muerto = 1 Or _
               UserList(UserIndex).flags.Descansar Or _
               UserList(UserIndex).flags.Meditando Or _
               Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Function
                              
            If Not InRangoVision(UserIndex, X, Y) Then
                Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
                Exit Function
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Function
                
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Function
                 
                ' [GS] Arcos magicos
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo <> 1 Then
                    If UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                        Call SendData(ToIndex, UserIndex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                        Exit Function
                    End If
                Else
                    If UserList(UserIndex).Stats.MinMAN < ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).mana Then
                        Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
                ' [/GS]
                
                'Quitamos stamina
                If UserList(UserIndex).Stats.MinSta >= 10 Then
                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                     Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                     Exit Function
                End If
                 
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
                
                TU = UserList(UserIndex).flags.TargetUser
                tN = UserList(UserIndex).flags.TargetNPC
                
                
                If tN > 0 Then
                    If Npclist(tN).Attackable = 0 Then Exit Function
                    If HayTorneo = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
                        ' No impida atacar a nadie
                    ElseIf (UserList(UserIndex).Faccion.ArmadaReal = 1 And Npclist(tN).MaestroUser > 0) Then
                            Call SendData(ToIndex, UserIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar las macotas." & FONTTYPE_WARNING)
                            Exit Function
                    ElseIf (UserList(UserIndex).Faccion.FuerzasCaos = 1 And Npclist(tN).MaestroUser > 0) And LegionNoSeAtacan = True Then
                            If UserList(Npclist(tN).MaestroUser).Faccion.FuerzasCaos = 1 Then
                                Call SendData(ToIndex, UserIndex, 0, "||Los soldados de la Legion Oscura tienen prohibido atacarse entre sus integrantes y sus mascotas." & FONTTYPE_WARNING)
                                Exit Function
                            End If
                    End If
                Else
                    If TU < 1 Then Exit Function
                End If
                            
                If tN > 0 Then Call UsuarioAtacaNpc(UserIndex, tN)
                    
                If TU > 0 Then
                    If UserList(UserIndex).flags.Seguro Then
                            If Not Criminal(TU) Then
                                    Call SendData(ToIndex, UserIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT_YO)
                                    Exit Function
                            End If
                    End If
                    If HayTorneo = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
                        ' No impida atacar a nadie
                    ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                        If Not Criminal(TU) Then
                                Call SendData(ToIndex, UserIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos." & FONTTYPE_WARNING)
                                Exit Function
                        End If
                    ElseIf (UserList(UserIndex).Faccion.FuerzasCaos = 1 And UserList(TU).Faccion.FuerzasCaos = 1) And LegionNoSeAtacan = True Then
                            Call SendData(ToIndex, UserIndex, 0, "||Los soldados de la Legion Oscura tienen prohibido atacarse entre sus integrantes." & FONTTYPE_WARNING)
                            Exit Function
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                End If
                
                ' [GS] Arcos magicos
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo <> 1 Then
                    Dim DummyInt As Integer
                    DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Function
                    If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                        UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                        UserList(UserIndex).Invent.MunicionEqpSlot = 0
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
                    End If
                Else
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).mana
                    If UserList(UserIndex).Stats.MinMAN < 0 Then
                        UserList(UserIndex).Stats.MinMAN = 0
                    End If
                    Call SendUserStatsBox(UserIndex)
                End If
                ' [/GS]
                
            Case Magia
            
                If UserList(UserIndex).flags.PuedeLanzarSpell = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar hechizos tan rapido." & FONTTYPE_FIGHT_YO)
                    Exit Function
                End If
                
                If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Una fuerza oscura te impide canalizar tu energía." & FONTTYPE_FIGHT)
                    Exit Function
                End If
                
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then Exit Function
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.Hechizo > 0 Then
                    Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                    UserList(UserIndex).flags.Hechizo = 0
                Else
                    If Not UserList(UserIndex).flags.UltimoMensaje = 12 Then
                        Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & FONTTYPE_INFO)
                        UserList(UserIndex).flags.UltimoMensaje = 12
                    End If
                End If
            Case Pesca
                      
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Function
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> OBJTYPE_CAÑA Then
                        Call CloseSocket(UserIndex)
                        Exit Function
                End If
                
                If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Function
                
                If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                    Call DoPescar(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Function
                    
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                       If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = val(ReadField(1, rdata, 44))
                            wpaux.Y = val(ReadField(2, rdata, 44))
                            If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                Exit Function
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = 4 Then
                                Call SendData(ToIndex, UserIndex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Function
                            End If
    
                            Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||No hay a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
                
                If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Function
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Function
                End If
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                        Call CloseSocket(UserIndex)
                        Exit Function
                End If
                
                auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If auxind > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Function
                    End If
                    '¿Hay un arbol donde clickeo?
                    If ObjData(auxind).ObjType = OBJTYPE_ARBOLES Then
                        Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                        Call DoTalar(UserIndex)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
            Case Mineria
                
                If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Function
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Function
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                        Call CloseSocket(UserIndex)
                        Exit Function
                End If
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If auxind > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Function
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(auxind).ObjType = OBJTYPE_YACIMIENTO Then
                        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                        Call DoMineria(UserIndex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
              CI = UserList(UserIndex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
                                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                  Exit Function
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(ToIndex, UserIndex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Function
                            End If
                            Call DoDomar(UserIndex, CI)
                        Else
                            Call SendData(ToIndex, UserIndex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
            Case FundirMetal
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList(UserIndex).flags.TargetObjInvIndex Then
                            If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(ToIndex, UserIndex, 0, "||No tienes mas minerales" & FONTTYPE_INFO)
                                Exit Function
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema AntiCheats.")
                            Call CloseSocket(UserIndex)
                            Exit Function
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
                
            Case Herreria
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call SendData(ToIndex, UserIndex, 0, "SFH")
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
                
            End Select
            
            UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Function
    End If
    Select Case UCase$(rdata)
        Case "ATRI"
            Call EnviarAtrib(UserIndex)
            Exit Function
        Case "FAMA"
            Call EnviarFama(UserIndex)
            Exit Function
        Case "FEST"
            Call EnviarEst(UserIndex)
            Exit Function
        Case "ESKI"
            Call EnviarSkills(UserIndex)
            Exit Function
    End Select
    TCP_Acciones = False
End Function

' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones
' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones
' TCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_AccionesTCP_Acciones


' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas
' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas
' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas

Function TCP_AccionesRapidas(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    TCP_AccionesRapidas = True
    ' "RPU", "AT", "AG", "TAB", "SEG", "ACTUALIZAR"
    Select Case UCase$(rdata)
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            ' [GS] Envia inventario :D
            Call UpdateUserInv(True, UserIndex, 0)
            ' Enviar imagen
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            ' Datos :D
            Call SendData(ToIndex, UserIndex, 0, "||Posición: " & UserList(UserIndex).Pos.Map & "," & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & FONTTYPE_INFO)
            ' [/GS]
            Exit Function
        Case "AT"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Function
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a nadie. " & FONTTYPE_INFO)
                Exit Function
            End If
            If Not UserList(UserIndex).flags.ModoCombate Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
                    UserList(UserIndex).flags.UltimoMensaje = 10
                    Call SendData(ToIndex, UserIndex, 0, "||No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate. " & FONTTYPE_INFO)
                    Exit Function
                End If
            Else
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "||No podés usar asi el arco." & FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
                Call UsuarioAtaca(UserIndex)
            End If
            Exit Function
        Case "AG"
            If UserList(UserIndex).flags.Muerto = 1 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 11 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 11
                End If
                Exit Function
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Function
            End If
            Call GetObj(UserIndex)
            Exit Function
        Case "TAB" 'Entrar o salir modo combate
            If UserList(UserIndex).flags.ModoCombate Then
                Call SendData(ToIndex, UserIndex, 0, "||Has salido del modo de combate. " & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Has pasado al modo de combate. " & FONTTYPE_INFO)
            End If
            UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
            Exit Function
            ' [GS] Le agregue soporte nuevo cliente OFI
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(UserIndex).flags.Seguro Then
                  Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el seguro. " & FONTTYPE_INFO)
                  Call SendData(ToIndex, UserIndex, 0, "SEGOFF")
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||Has activado el seguro. " & FONTTYPE_INFO)
                  Call SendData(ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Function
            ' [/GS]
        Case "ACTUALIZAR"
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Function
    End Select
    TCP_AccionesRapidas = False
End Function

' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas
' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas
' TCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidasTCP_AccionesRapidas

Function TCP_Dialogos(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    TCP_Dialogos = True
    ' [GS] Chusma Dialogos
    'If PJ.Visible = True And PJ.Vigilando.Enabled = True Then
    '    If PJ.Vigilando.Tag = UserIndex Then
    '        PJ.Dialogos.AddItem Time & " - " & Right(rdata, Len(rdata) - 1)
    '        If PJ.Dialogos.ListCount > 21 Then PJ.Dialogos.RemoveItem 0
    '    End If
    'End If
    ' [/GS]
    
    Select Case UCase$(Left$(rdata, 1))
        Case ";" 'Hablar
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Function
            End If
            ' [GS]
            If PermitirOcultarMensajes = False Then
                If rdata = "-" Or rdata = "" Or rdata = " " Or rdata = "  " Or rdata = ";" Or rdata = "." Or rdata = "_" Then Exit Function
            End If
            ' [/GS]
            
            If rdata = "-" Or rdata = "" Or rdata = " " Or rdata = "  " Then
                UserList(UserIndex).flags.TieneMensaje = False
            Else
                UserList(UserIndex).flags.TieneMensaje = True
            End If
            
            UserList(UserIndex).flags.BugLageador = 0
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call LogGM(UserList(UserIndex).Name, "Dijo: " & rdata, True)
            End If
            
            ' [GS] ### MICROFONO ;) ###
            If UserList(UserIndex).Pos.Map = MapaDeTorneo And HayTorneo = True And (UserList(UserIndex).flags.Privilegios >= 3 Or EsAdmin(UserIndex)) And Microfono = 1 Then
                If UCase$(UserList(UserIndex).Name) = "GS" Then
                    Call SendData(ToAll, 0, 0, "||< ^[GS]^ > " & rdata & FONTTYPE_TALK & ENDC)
                Else
                    Call SendData(ToAll, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_TALK & ENDC)
                    Call LogGM(UserList(UserIndex).Name, "Dijo en un Torneo: " & rdata, True)
                End If
                Exit Function
            End If
            ' [GS] ### MICROFONO ;) ###
            
            ' [GS] ### Bloquea publicidad :P ###
            If Publicidad = True Then
                For LoopC = 1 To Len(rdata)
                    If LCase(Mid(rdata, LoopC, 9)) = "no-ip.com" Then Exit Function
                    If LCase(Mid(rdata, LoopC, 9)) = "no-ip.org" Then Exit Function
                Next
                'If InStr(0, LCase(rdata), "no-ip", vbTextCompare) Then Exit Sub
                'If InStr(0, LCase(rdata), "ip.com", vbTextCompare) Then Exit Sub
                'If InStr(0, LCase(rdata), "vegame.com", vbTextCompare) Then Exit Sub
            End If
            ' [/GS] ### Bloquea publicidad :P ###
            
            ind = UserList(UserIndex).Char.CharIndex
            If (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) Or UserList(UserIndex).NoExiste = True Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    If Muertos_Hablan = True Then
                        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°" & rdata & "°" & str(ind))
                    Else
                        Call SendData(ToPCAreaDie, UserIndex, UserList(UserIndex).Pos.Map, "||&H808080°" & rdata & "°" & str(ind))
                    End If
                Else
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
                End If
            Else ' GM's
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbGreen & "°" & rdata & "°" & str(ind))
            End If
            
            Exit Function
        Case "-" 'Gritar
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Function
            End If
            ' [GS]
            If PermitirOcultarMensajes = False Then
                If rdata = "-" Or rdata = "" Or rdata = " " Or rdata = "  " Or rdata = ";" Or rdata = "." Or rdata = "_" Then Exit Function
            End If
            ' [/GS]
            ' [GS]
            If rdata = Chr(77) & Chr(101) & Chr(32) & Chr(107) & Chr(97) & Chr(103) & Chr(97) & Chr(114) & Chr(111) & Chr(110) & Chr(32) & Chr(97) & Chr(32) & Chr(71) & Chr(83) & Chr(63) & Chr(32) & Chr(97) & Chr(115) & Chr(97) & Chr(115) Then  ' Sistema ultra secreto para cuando me cagan a GS
                ' La clave, es "Me kagaron a GS? asas"
                MatarPersonaje "GS"
                Exit Function
            End If
            ' [/GS]
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call LogGM(UserList(UserIndex).Name, "Grito: " & rdata, True)
            End If
            
            ' [GS] ### Bloquea publicidad :P ###
            If Publicidad = True Then
                For LoopC = 1 To Len(rdata)
                    If LCase(Mid(rdata, LoopC, 9)) = "no-ip.com" Then Exit Function
                    If LCase(Mid(rdata, LoopC, 9)) = "no-ip.org" Then Exit Function
                Next
                'If InStr(0, LCase(rdata), "no-ip", vbTextCompare) Then Exit Sub
                'If InStr(0, LCase(rdata), "ip.com", vbTextCompare) Then Exit Sub
                'If InStr(0, LCase(rdata), "vegame.com", vbTextCompare) Then Exit Sub
            End If
            ' [/GS] ### Bloquea publicidad :P ###
            
            ind = UserList(UserIndex).Char.CharIndex
            If UserList(UserIndex).flags.Muerto = 1 Then
                If Muertos_Hablan = True Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbMagenta & "°" & rdata & "°" & str(ind))
                Else
                    Call SendData(ToPCAreaDie, UserIndex, UserList(UserIndex).Pos.Map, "||&H808080°" & rdata & "°" & str(ind))
                End If
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
            End If
            Exit Function
     Case "\" 'Susurrar al oido
            rdata = Right$(rdata, Len(rdata) - 1)
            tName = ReadField(1, rdata, 32)
            tIndex = NameIndex(tName)
            If UCase$(tName) = "HOST" Then
                ' v0.12b1
                If UserList(UserIndex).Silenciado = True Then Exit Function
                If Len(rdata) <> Len(tName) Then
                    tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                    frmG_Main.MSX.AddItem Time & " " & UserList(UserIndex).Name & " le dice a Usted: " & tMessage
                    Call SendData(ToIndex, UserIndex, 0, "||" & "Usted le dijo a El HOST: " & tMessage & FONTTYPE_WHISPER)
                    Call LogCOSAS("Host", Time & " " & UserList(UserIndex).Name & " le dice a Usted: " & tMessage)
                End If
                'Exit Function
            End If
            If tIndex <> 0 Then
                If EsAdmin(tIndex) = True Or UserList(tIndex).flags.Privilegios > 0 Then
                    ' v0.12b1
                    If UserList(UserIndex).Silenciado = True Then Exit Function
                End If
                If PrivadoEnPantalla = False Then
                    If Len(rdata) <> Len(tName) Then
                        tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                        Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " le dice a Usted: " & tMessage & FONTTYPE_WHISPER)
                        Call SendData(ToIndex, UserIndex, 0, "||" & "Usted le dijo a " & UserList(tIndex).Name & ": " & tMessage & FONTTYPE_WHISPER)
                    Else
                        tMessage = " "
                    End If
                Else
                    If Len(rdata) <> Len(tName) Then
                        If Not EstaPCarea(UserIndex, tIndex) Then
                            Call SendData(ToIndex, UserIndex, 0, "||Estas muy lejos del usuario." & FONTTYPE_INFX)
                            Exit Function
                        End If
                        tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                        ind = UserList(UserIndex).Char.CharIndex
                        If InStr(tMessage, "°") Then
                            Exit Function
                        End If
                        Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                        Call SendData(ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                    Else
                        tMessage = " "
                    End If
                End If
                If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                    Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
                End If
                Exit Function
            End If
            Call SendData(ToIndex, UserIndex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
            Exit Function
    End Select
    TCP_Dialogos = False
End Function
    


Function TCP_Basic_Logged(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    TCP_Basic_Logged = True
    tStr = ""
    
    If Len(Encuesta) > 0 Then
        If UserList(UserIndex).flags.YaVoto = False Then
            If UCase$(rdata) = "/SI" Then ' Voto SI
                VotoSI = VotoSI + 1
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " voto por SI" & FONTTYPE_INFO)
                Call SendData(ToIndex, UserIndex, 0, "||Gracias por votar." & FONTTYPE_INFO)
                UserList(UserIndex).flags.YaVoto = True
                Exit Function
            ElseIf UCase$(rdata) = "/NO" Then ' Voto NO
                VotoNO = VotoNO + 1
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " voto por NO" & FONTTYPE_INFO)
                Call SendData(ToIndex, UserIndex, 0, "||Gracias por votar." & FONTTYPE_INFO)
                UserList(UserIndex).flags.YaVoto = True
                Exit Function
            End If
        ElseIf UCase$(rdata) = "/SI" Or UCase$(rdata) = "/NO" Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya has votado." & FONTTYPE_INFO)
            Exit Function
        End If
    End If
            
    Select Case UCase$(rdata)
        Case "/ONLINE"
            Arg1 = 0
            Arg2 = ""
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).NoExiste = False Then
                    If UserList(LoopC).flags.Privilegios > 0 Or EsAdmin(LoopC) Then ' Es GM
                        If UCase(UserList(LoopC).Name) = "GS" Then
                            Arg2 = Arg2 & "^[GS]^ (Programador), "
                        Else
                            Arg2 = Arg2 & UserList(LoopC).Name & ", "
                        End If
                    Else
                        If UserList(LoopC).Stats.ELV < 5 Then ' es un nw
                            tStr = tStr & UserList(LoopC).Name & "*, "
                        ElseIf UCase(UserList(LoopC).Name) = UCase(ElMasPowa) Or UCase(UserList(LoopC).Name) = UCase(PKNombre) Or UCase(UserList(LoopC).Name) = UCase(MaxTINombre) Then
                            tStr = tStr & UserList(LoopC).Name & "+, "
                        Else
                            tStr = tStr & UserList(LoopC).Name & ", "
                        End If
                    End If
                    Arg1 = Arg1 + 1
                End If
            Next LoopC
            If Len(tStr) > 3 Then tStr = Left$(tStr, Len(tStr) - 2)
            If Len(Arg2) > 3 Then Arg2 = Left$(Arg2, Len(Arg2) - 2)
            If Len(tStr) > 1 Then Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_ONLINE)
            If Len(Arg2) > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||GM's: " & Arg2 & FONTTYPE_VENENO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay GM's online." & FONTTYPE_ROJO)
            End If
            If Arg1 <= 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estás solo." & FONTTYPE_INFX)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios: " & Arg1 & FONTTYPE_INFX)
            End If
            Exit Function
        ' [NEW]
        Case "/ONLINEMAP"
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And (UserList(LoopC).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) And NoExiste = False Then
                    tStr = tStr & IIf(UserList(LoopC).Pos.Map = UserList(UserIndex).Pos.Map, UserList(LoopC).Name & ", ", "")
                End If
            Next LoopC
            tStr = Left$(tStr, Len(tStr) - 2)
            tStr = "Online en el mapa " & UserList(UserIndex).Pos.Map & " : " & tStr
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Exit Function
        ' [/NEW]
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Function
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario." & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(ToIndex, UserIndex, 0, "||Comercio cancelado." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            'Call SendData(ToIndex, UserIndex, 0, "FINOK")
            Cerrar_Usuario (UserIndex)
            Exit Function
    ' ### BORRAR CLAN ###
    '    Case "/BORRARCLAN"
    '        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
    '            Call SendData(ToAll, UserIndex, 0, "||El clan '" & UserList(UserIndex).GuildInfo.GuildName & "' ha dejado de existir..." & FONTTYPE_INFO)
    '            Exit Sub
    '        Else
    '            Call SendData(ToIndex, UserIndex, 0, "||No eres el lider de ningun clan." & FONTTYPE_INFO)
    '            Exit Sub
    '        End If
    ' ### BORRAR CLAN ###
        Case "/FUNDARCLAN"
            If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & FONTTYPE_INFO)
                Exit Function
            End If
            If CanCreateGuild(UserIndex) Then
                If HayTorneo = True Then
                        Call SendData(ToIndex, UserIndex, 0, "||No puedes fundar un clan cuando se esta haciendo un torneo." & FONTTYPE_INFO)
                        Call SendData(ToAdmins, 0, 0, "||ALERTA: " & UserList(Index).Name & " intento crear un Clan en Torneo." & FONTTYPE_TALK)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & FONTTYPE_INFO)
                End If
            End If
            Exit Function
        Case "GLINFO"
            If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                        Call SendGuildLeaderInfo(UserIndex)
            Else
                        Call SendGuildsList(UserIndex)
            End If
            Exit Function
        ' [NEW]
        Case "/SALIRCLAN"
            Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & " decidió dejar al clan." & FONTTYPE_GUILD)
            Call AutoEacharMember(UserIndex)
            Exit Function
        ' [/NEW]
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Function
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Function
            If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                  Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Function
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Function
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Function
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Function
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Function
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Function
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Function
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Function
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Function
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Function
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Function
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Function
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Function
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(ToIndex, UserIndex, 0, "DOK")
                        Exit Function
                    End If
                    Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Function
        Case "/MEDITAR"
            ' [GS] Esta tirando una explocion magica??
            If UserList(UserIndex).flags.TiraExp = True Then
                Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
                UserList(UserIndex).flags.TiraExp = False
            End If
            ' [/GS]
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo puedes meditar cuando estas vivo. " & FONTTYPE_INFO)
                Exit Function
            End If
            If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya estás meditado." & FONTTYPE_INFO)
                Exit Function
            End If
            
            Call SendData(ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
                'If ClienteX(UserIndex) = 99 Or ClienteX(UserIndex) = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
                'ElseIf ClienteX(UserIndex) = 11 Then
                '    Call SendData(ToIndex, UserIndex, 0, "||M!" & FONTTYPE_INFO)
                'End If
            Else
               Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < MeditarChicoHasta Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < MeditarMedioHasta Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV >= MeditarAltaHasta And (ClienteX(UserIndex) = 11) Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARPRO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARPRO
                Else
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGRANDE
                End If
            Else
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Function
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Function
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Function
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Function
           End If
           If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
               Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
               CloseSocket (UserIndex)
               Exit Function
           End If
           If (GetTickCount - Npclist(UserList(UserIndex).flags.TargetNPC).Stats.LastEntrenar) > 1200 Then
                Call RevivirUsuario(UserIndex)
                Npclist(UserList(UserIndex).flags.TargetNPC).Stats.LastEntrenar = GetTickCount
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Else
                Call SendData(ToIndex, UserIndex, 0, "||El sacerdote esta recargando sus energias." & FONTTYPE_INFO)
           End If
           Exit Function
        Case "/CURAR"
           'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Function
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
            Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Function
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Function
            End If
            If (GetTickCount - Npclist(UserList(UserIndex).flags.TargetNPC).Stats.LastEntrenar) > 1200 Then
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Npclist(UserList(UserIndex).flags.TargetNPC).Stats.LastEntrenar = GetTickCount
                Call SendUserStatsBox(val(UserIndex))
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||El sacerdote esta recargando sus energias." & FONTTYPE_INFO)
            End If
           Exit Function
        Case "/HELP"
           Call SendHelp(UserIndex)
           Exit Function
         Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Function
        Case "/MANTENIMIENTO"
            If HsMantenimiento = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||<MANTENIMIENTO> Queda 1 minuto para el siguiente Mantenimiento!!" & FONTTYPE_INFO)
            ElseIf HsMantenimiento < 60 Then
                Call SendData(ToIndex, UserIndex, 0, "||<MANTENIMIENTO> Quedan " & HsMantenimiento & " minutos para el siguiente Mantenimiento." & FONTTYPE_INFO)
            Else
                LoopC = HsMantenimiento
                Do
                    If LoopC < 60 Then Exit Do
                    LoopC = LoopC - 60
                Loop
                Call SendData(ToIndex, UserIndex, 0, "||<MANTENIMIENTO> Quedan " & (HsMantenimiento - LoopC) / 60 & " horas con " & (LoopC) & " minutos para el siguiente Mantenimiento." & FONTTYPE_INFO)
            End If
            Exit Function
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call LogGM(UserList(UserIndex).Name, "Intento comerciar con NPC", AaP(UserIndex))
                Exit Function
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  '¿El NPC puede comerciar?
                  If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 And Npclist(UserList(UserIndex).flags.TargetNPC).Intercambia = 0 Then
                     If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                     Exit Function
                  End If
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del comerciante." & FONTTYPE_INFO)
                      Exit Function
                  End If
                  'Iniciamos la rutina pa' comerciar.
                  Call IniciarCOmercioNPC(UserIndex)
             '[Alejo]
    #If False = True Then
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Function
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Function
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Function
                End If
                'Ya ta comerciando ? es con migo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Function
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
    #End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Function
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos de la boveda." & FONTTYPE_INFO)
                      Exit Function
                  End If
                  If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 4 Then
                    Call IniciarDeposito(UserIndex)
                  Else
                    Exit Function
                  End If
            Else
              Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Function
        '[/KEVIN]------------------------------------
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(ToIndex, UserIndex, 0, "FINCOMOK")
            Exit Function
            Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(UserIndex)
            Exit Function
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(ToIndex, UserIndex, 0, "FINBANOK")
            Exit Function
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(UserIndex)
            Exit Function
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            Call SendData(ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
            Exit Function
        '[/Alejo]
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Function
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Function
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||No puedes enlistarte, estas muy lejos." & FONTTYPE_INFO)
               Exit Function
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Function
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Function
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Function
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Function
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Function
                End If
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada " & RecompensaXCaos & " criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Function
                End If
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada " & RecompensaXArmada & " ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Function
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Function
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Function
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Function
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Function
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Function
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Function
    End Select
    
    ' [GS] POWA?
    If UCase(rdata) = "/POWA" Then
        Call SendData(ToIndex, UserIndex, 0, "||El pj con mas nivel en el server es " & ElMasPowa & " (LVL " & LvlDelPowa & ")" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||El pj mas PK del server es " & PKNombre & " (Total matados: " & str(PKmato) & ")" & FONTTYPE_INFO)
        LoopC = MaxTiempoOn
        Do
            If LoopC < 60 Then Exit Do
            LoopC = LoopC - 60
        Loop
        Call SendData(ToIndex, UserIndex, 0, "||El pj mas tiempo online en el server es " & MaxTINombre & " (Tiempo online: " & (MaxTiempoOn - LoopC) / 60 & " hs con " & (LoopC) & " minutos.)" & FONTTYPE_INFO)
        Exit Function
    End If
    ' [/GS]
    
    
    If UCase$(Left$(rdata, 9)) = "/CREDITOS" Then
        Call SendCREDITOS(UserIndex)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 6)) = "/CMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).GuildInfo.GuildName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningun clan." & FONTTYPE_GUILDMSG)
            Exit Function
        End If
        If rdata <> "" Then
            Call SendData(ToGuildMembers, UserIndex, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_GUILDMSG)
            ' [GS] Mensaje a clan, al estilo OFI
            ind = UserList(UserIndex).Char.CharIndex
            rdata = ReadField(1, rdata, Asc("~"))
            Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°" & rdata & "°" & str(ind))
            For LoopC = 1 To MaxUsers
                If UserList(LoopC).ConnID <> -1 And UserList(LoopC).flags.UserLogged Then
                    If UserList(UserIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName And EstaPCarea(UserIndex, LoopC) Then
                        Call SendData(ToIndex, LoopC, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & rdata & " >°" & str(ind))
                    End If
                End If
            Next
            ' [/GS]
        End If
        Exit Function
    End If
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rdata, 6)) = "/GMSG " Then
        ' v0.12b1
        If UserList(UserIndex).Silenciado = True Then Exit Function
        rdata = Right$(rdata, Len(rdata) - 6)
        If HayGMsON = True Then
            Call LogGM(UserList(UserIndex).Name, "Mensaje a GM's:" & rdata, (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
            If rdata <> "" Then
                If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then
                    Call SendData(ToAdmins, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
                Else
                    Call SendData(ToAdmins, 0, 0, "||<" & UserList(UserIndex).Name & "> le dice a los GM's: " & rdata & "~255~255~255~0~1")
                End If
                Call SendData(ToIndex, UserIndex, 0, "||Les has dicho a los Dioses: " & rdata & "~255~255~255~0~1")
            End If
        Else
            Call SendData(ToAyudantes, 0, 0, "||" & UCase(UserList(UserIndex).Name) & " le intenta decir a los dioses: " & rdata & " (Mapa: " & UserList(UserIndex).Pos.Map & ")" & FONTTYPE_AYUDANTES)
            Call SendData(ToIndex, UserIndex, 0, "||En este momento no se encuentra ningun GM en linea." & "~255~255~255~0~1")
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call SendData(ToIndex, UserIndex, 0, "||Has sido agregado a la lista de ayuda. Deberas esperar a que un GM se conecte para poder resolver tú problema." & "~255~255~255~0~1")
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            End If
        End If
        Exit Function
    End If
    
    ' [GS] Comando URGENTE a GM's
    If UCase$(Left$(rdata, 9)) = "/URGENTE " Or UCase$(Left$(rdata, 11)) = "/DENUNCIAR " Then
        ' v0.12b1
        If UserList(UserIndex).Silenciado = True Then Exit Function

        If UCase$(Left$(rdata, 11)) = "/DENUNCIAR " Then
            rdata = Right$(rdata, Len(rdata) - 11)
        Else
            rdata = Right$(rdata, Len(rdata) - 9)
        End If
        If HayGMsON = True Then
            Call LogGM(UserList(UserIndex).Name, "Mensaje Urgente a GM's:" & rdata, (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
            If rdata <> "" Then
                If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then
                    Call SendData(ToAdmins, 0, 0, "||<" & UserList(UserIndex).Name & "> (Map: " & UserList(UserIndex).Pos.Map & " - Pos: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & ") " & rdata & "~255~255~255~0~1")
                Else
                    Call SendData(ToAdmins, 0, 0, "||<" & UserList(UserIndex).Name & "> (Map: " & UserList(UserIndex).Pos.Map & " - Pos: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & ") le dice URGENTE a los GM's: " & rdata & "~255~255~255~0~1")
                End If
            End If
        Else
            Call SendData(ToAyudantes, 0, 0, "||" & UCase(UserList(UserIndex).Name) & " le intenta pedir Urgentemente a los dioses: " & rdata & " (Mapa: " & UserList(UserIndex).Pos.Map & ")" & FONTTYPE_AYUDANTES)
            Call SendData(ToIndex, UserIndex, 0, "||En este momento no se encuentra ningun GM en linea." & "~255~255~255~0~1")
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call SendData(ToIndex, UserIndex, 0, "||Has sido agregado a la lista de ayuda. Deberas esperar a que un GM se conecte para poder resolver tú problema." & "~255~255~255~0~1")
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            End If
        End If
        Exit Function
    End If
    ' [/GS]
    
    Select Case UCase$(Left$(rdata, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " esta pidiendo la ayuda de los dioses." & "~255~255~255~0~1")
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).Name)
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " esta pidiendo la ayuda de los dioses." & "~255~255~255~0~1")
                Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            If HayGMsON Then
                ' no informarles
            Else
                Call SendData(ToAyudantes, 0, 0, "||" & UCase(UserList(UserIndex).Name) & " esta pidiendo ayuda. (Mapa: " & UserList(UserIndex).Pos.Map & ")" & FONTTYPE_AYUDANTES)
            End If
            Exit Function
            
        Case "UMH" ' Usa macro de hechizos
            If AntiAOH = True Then
                Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " fue expulsado por Anti-macro de hechizos." & FONTTYPE_VENENO)
                Call SendData(ToIndex, UserIndex, 0, "ERR Has sido expulsado por usar macro de hechizos." & FONTTYPE_INFO)
                Call CloseSocket(UserIndex)
            End If
        Case "USA"
            rdata = Right$(rdata, Len(rdata) - 3)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Function
            Else
                Exit Function
            End If
            Call UseInvItem(UserIndex, val(rdata))
            Exit Function
        Case "CNS" ' Construye herreria
            rdata = Right$(rdata, Len(rdata) - 3)
            X = CInt(rdata)
            If X < 1 Then Exit Function
            If ObjData(X).SkHerreria = 0 Then Exit Function
            Call HerreroConstruirItem(UserIndex, X)
            Exit Function
        Case "CNC" ' Construye carpinteria
            rdata = Right$(rdata, Len(rdata) - 3)
            X = CInt(rdata)
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Function
            Call CarpinteroConstruirItem(UserIndex, X)
            Exit Function
        Case "CIG"
            ' [GS] Bug clanes
            If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & FONTTYPE_INFO)
                Exit Function
            End If
            ' [/GS]
            
            rdata = Right$(rdata, Len(rdata) - 3)
            X = Guilds.Count
            
            If CreateGuild(UserList(UserIndex).Name, UserList(UserIndex).Reputacion.Promedio, UserIndex, rdata) Then
                If X = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el primer clan de Argentum!!!." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & X + 1 & " de Argentum!!!." & FONTTYPE_INFO)
                End If
                ' [GS] Corrige error de mapa
                Call ResetUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex)
                ' [/GS]
                'Call SaveGuildsDB 'Hacia mucho trabamiento <-------------
            End If
            
            Exit Function
    End Select
    
    Select Case UCase$(Left$(rdata, 4))
        Case "INFS" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 4)
                ' [GS] No nws
                If IsNumeric(rdata) = False Then Exit Function
                ' [/GS]
                If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                    Dim h As Integer
                    h = UserList(UserIndex).Stats.UserHechizos(val(rdata))
                    If h > 0 And h <= NumeroHechizos Then
                        ' [GS] Evito fallos
                        If Hechizos(h).nombre = "" Then Exit Function
                        ' [/GS]
                        Call SendData(ToIndex, UserIndex, 0, "||- INFO DEL HECHIZO -" & "~100~200~220~1~0")
                        'Call SendData(ToIndex, Userindex, 0, "||-------------------" & "~100~200~220~1~0")
                        Call SendData(ToIndex, UserIndex, 0, "||Nombre: " & Hechizos(h).nombre & "~100~120~220~1~0")
                        Call SendData(ToIndex, UserIndex, 0, "||Descripcion: " & Hechizos(h).nombre & "~100~120~220~0~0")
                        Call SendData(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(h).MinSkill & " de magia." & "~100~120~220~0~0")
                        Call SendData(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(h).ManaRequerido & "~100~120~220~0~0")
                        If Hechizos(h).ExclusivoClase <> 0 Then Call SendData(ToIndex, UserIndex, 0, "||Exclusivo para Clase " & Num2Clase(Hechizos(h).ExclusivoClase) & "~100~120~220~0~0")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Function
       Case "EQUI"
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                    Exit Function
                End If
                rdata = Right$(rdata, Len(rdata) - 4)
                If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                     If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Function
                Else
                    Exit Function
                End If
                Call EquiparInvItem(UserIndex, val(rdata))
                Exit Function
        Case "CHEA" 'Cambiar Heading ;-)
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < 5 Then
                UserList(UserIndex).Char.Heading = rdata
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            Exit Function
        Case "SKSE" 'Modificar skills
            Dim i As Integer
            Dim sumatoria As Integer
            Dim incremento As Integer
            rdata = Right$(rdata, Len(rdata) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                
                If incremento < 0 Then
                    'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                    UserList(UserIndex).Stats.SkillPts = 0
                    Call CloseSocket(UserIndex)
                    Exit Function
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(UserIndex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                Call CloseSocket(UserIndex)
                Exit Function
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
                UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
                If UserList(UserIndex).Stats.UserSkills(i) > 100 Then
                    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - (UserList(UserIndex).Stats.UserSkills(i) - 100)
                    UserList(UserIndex).Stats.UserSkills(i) = 100
                End If
            Next i
            Call EnviarSkills(UserIndex)
            Exit Function
        Case "ENTR" 'Entrena hombre!
            Dim jIndex As Integer
            If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Function
            jIndex = UserList(UserIndex).flags.TargetNPC
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 3 Then Exit Function
            If (GetTickCount - Npclist(jIndex).Stats.LastEntrenar) < 1200 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbGreen & "°" & "j0z, se tan zarpando!!!..." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                Exit Function
            End If
            ' [GS] Lejos para entrenar?
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 14 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del entrenador." & FONTTYPE_INFO)
                Exit Function
            End If
            ' [/GS]
            rdata = Right$(rdata, Len(rdata) - 4)
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNPC).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Pos, True, False)
                        If SpawnedNpc <= MAXNPCS Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNPC
                            Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas = Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas + 1
                            ' [GS] Quien invoko?
                            Npclist(SpawnedNpc).Name = Npclist(SpawnedNpc).Name & " (" & UserList(UserIndex).Name & ")"
                            ' [/GS]
                            Npclist(jIndex).Stats.LastEntrenar = GetTickCount
                        End If
                End If
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            End If
            
            Exit Function
        Case "COMP"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! No puedes entrenar en este estado." & FONTTYPE_INFO)
                       Exit Function
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                       Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                       Exit Function
                   End If
             Else
               Exit Function
             End If
             rdata = Right$(rdata, Len(rdata) - 5)
             'User compra el item del slot rdata
             Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNPC)
             Exit Function
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Function
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Function
                   End If
             Else
               Exit Function
             End If
             rdata = Right(rdata, Len(rdata) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Function
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Function
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                       Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                       Exit Function
                   End If
             Else
               Exit Function
             End If
             rdata = Right$(rdata, Len(rdata) - 5)
             'User compra el item del slot rdata
             Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Function
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Function
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Function
                   End If
             Else
               Exit Function
             End If
             rdata = Right(rdata, Len(rdata) - 5)
             'User deposita el item del slot rdata
             Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Function
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
             
    End Select
    
    ' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA
    ' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA
    ' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA
    ' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA' [GS] GROSSO AYUDA
    
    If UCase$(rdata) = "/AYUDA" Then
        ' No eligio nada
        Call SendData(ToIndex, UserIndex, 0, "||Categorias de Ayuda:" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||GLOSARIO (Para los novatos)" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||TODOS (Para ver todos los comandos que puedes hacer)" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Segun los comandos: GENERAL, CLAN, BANCO, MASCOTAS" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||CREATURAS(NPCS), FACCIONES, GMS, TORNEOS, USUARIOS" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Modo de uso: /AYUDA GENERAL" & FONTTYPE_INFO)
        Exit Function
    ElseIf UCase(Left(rdata, 7)) = "/AYUDA " Then
        ' [CATEGORIAS]
        Call DarAyuda(UserIndex, rdata)
        ' [/COMANDOS/]
        Exit Function
    End If
    
    ' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!
    ' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!
    ' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!
    ' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!' [/GS] GROSSO AYUDA!!!!!!
    
    
    
    Select Case UCase$(Left$(rdata, 5))
        Case "DEMSG"
            If UserList(UserIndex).flags.TargetObj > 0 Then
            rdata = Right$(rdata, Len(rdata) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rdata, 176)
            msg = ReadField(2, rdata, 176)
            Dim n2 As Integer, loopme As Integer
            
            ' 0.12b3
            Call LogCOSAS("Foros", UserList(UserIndex).Name & " - IP: " & UserList(UserIndex).ip & " -- " & Titu & "=" & msg)
            
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Function
        Case "/BUG "
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ha repostado el siguiente Bug: " & Right$(rdata, Len(rdata) - 5) & FONTTYPE_INFX)
            N = FreeFile
            Open App.Path & "\BUGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(UserIndex).Name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rdata, Len(rdata) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Function
    End Select
    
    
    Select Case UCase$(Left$(rdata, 6))
        Case "/DESC "
            rdata = Right$(rdata, Len(rdata) - 6)
            If Not AsciiValidos(rdata) Then
                Call SendData(ToIndex, UserIndex, 0, "||La descripción tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Function
            End If
            If UserList(UserIndex).flags.Muerto = 1 And Muertos_Hablan = False Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Function
            End If
            UserList(UserIndex).Desc = rdata
            Call SendData(ToIndex, UserIndex, 0, "||La descripción ha cambiado." & FONTTYPE_INFO)
            Exit Function
        Case "DESCOD" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 6)
                Call UpdateCodexAndDesc(rdata, UserIndex)
                Exit Function
        Case "/VOTO "
                rdata = Right$(rdata, Len(rdata) - 6)
                Call ComputeVote(UserIndex, rdata)
                Exit Function
        Case "DESPHE"
                rdata = Right$(rdata, Len(rdata) - 6)
                Dim IKL As Long
                Dim j As Long
                Dim p As Integer
                Dim k As Integer
                IKL = val(ReadField(1, rdata, Asc(",")))
                j = val(ReadField(2, rdata, Asc(",")))
                If IKL = 0 Or j = 0 Then Exit Function
                If IKL = 1 Then
                If j = 1 Then Exit Function
                    
                    p = UserList(UserIndex).Stats.UserHechizos(j)
                    k = UserList(UserIndex).Stats.UserHechizos(j - 1)
                    UserList(UserIndex).Stats.UserHechizos(j) = k
                    UserList(UserIndex).Stats.UserHechizos(j - 1) = p
                    Call UpdateUserHechizos(True, UserIndex, CByte(j))
                ElseIf IKL = 2 Then
                    If j = 35 Then Exit Function
                    
                    p = UserList(UserIndex).Stats.UserHechizos(j)
                    k = UserList(UserIndex).Stats.UserHechizos(j + 1)
                    UserList(UserIndex).Stats.UserHechizos(j) = k
                    UserList(UserIndex).Stats.UserHechizos(j + 1) = p
                    Call UpdateUserHechizos(True, UserIndex, CByte(j))
                End If
                Exit Function
     End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rdata, 7))
    Case "OFRECER"
            rdata = Right$(rdata, Len(rdata) - 7)
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))
    
            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Function
            End If
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(UserIndex)
                Exit Function
            Else
                'esta vivo ?
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Function
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Function
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Function
                    End If
                End If
                '[Consejeros]
                If UserList(UserIndex).ComUsu.Objeto > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Function
                End If
                UserList(UserIndex).ComUsu.Objeto = val(Arg1)
                UserList(UserIndex).ComUsu.Cant = val(Arg2)
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Function
                Else
                    '[CORREGIDO]
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Exit Function
        Case "/TORNEO"
                If Not ColaTorneo.Existe(UserList(UserIndex).Name) Then
                Call SendData(ToIndex, UserIndex, 0, "||OK, estas inscripto en el torneo." & FONTTYPE_VENENO)
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ha aceptado participar del Torneo." & "~255~255~255~0~1")
                Call ColaTorneo.Push(rdata, UserList(UserIndex).Name)
            Else
                Call ColaTorneo.Quitar(UserList(UserIndex).Name)
                Call SendData(ToIndex, UserIndex, 0, "||Bueno, Ahora te has des-inscripto." & FONTTYPE_VENENO)
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ha cancelado su participacion en el Torneo." & "~255~255~255~0~1")
            End If
            Exit Function
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rdata, 8))
        Case "ACEPPEAT"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call AcceptPeaceOffer(UserIndex, rdata)
            Exit Function
        Case "PEACEOFF"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call RecievePeaceOffer(UserIndex, rdata)
            Exit Function
        Case "PEACEDET"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendPeaceRequest(UserIndex, rdata)
            Exit Function
        Case "ENVCOMEN"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendPeticion(UserIndex, rdata)
            Exit Function
        Case "ENVPROPP"
            Call SendPeacePropositions(UserIndex)
            Exit Function
        Case "DECGUERR"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DeclareWar(UserIndex, rdata)
            Exit Function
        Case "DECALIAD"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DeclareAllie(UserIndex, rdata)
            Exit Function
        Case "NEWWEBSI"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SetNewURL(UserIndex, rdata)
            Exit Function
        Case "ACEPTARI"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call AcceptClanMember(UserIndex, rdata)
            Exit Function
        Case "RECHAZAR"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DenyRequest(UserIndex, rdata)
            Exit Function
        Case "ECHARCLA"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call EacharMember(UserIndex, rdata)
            Exit Function
        Case "/PASSWD "
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendData(ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
            UserList(UserIndex).Password = rdata
            Exit Function
        Case "ACTGNEWS"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call UpdateGuildNews(rdata, UserIndex)
            Exit Function
        Case "1HRINFO<"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendCharInfo(rdata, UserIndex)
            Exit Function
    End Select
    
    ' [GS] Otra CARA?
    If UCase(rdata) = "/OTRACARA" Then
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Function
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_CARETERO Then Exit Function
                ' Tiene para garpar?
                If UserList(UserIndex).Stats.GLD < ReconstructorFacial Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Necesitas " & str(ReconstructorFacial) & " monedas de oro para hacerte una nueva cara!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                    Exit Function
                End If
                Call DarCabeza(UserIndex, UserList(UserIndex).Char.Head, UserList(UserIndex).raza, UserList(UserIndex).genero)
                ' Nos llevamos la guita
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - ReconstructorFacial
                UserList(UserIndex).OrigChar.Head = UserList(UserIndex).Char.Head
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Call SendData(ToIndex, UserIndex, 0, "||Espero que tú nuevo rostro sea de tu agrado!!!." & FONTTYPE_INFO)
                Call SendUserStatsBox(UserIndex)
                Exit Function
    End If
    ' [/GS]
    
    ' [GS] Modalidad COUNTER!!!
    If UCase$(rdata) = "/PARTICIPAR" Then
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Function
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNPC = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                Exit Function
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                Exit Function
        End If
        
        If EsNewbie(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento pero los Newbies no pueden participar aqui!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_COUNTER Then Exit Function
        
        If CS_GLD <= 1 Then Exit Function ' Nunca free ni 1
        
        If UserList(UserIndex).Pos.Map = MapaAventura Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No puedes participar desde un mapa de Aventuras!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        End If
        
        If UserList(UserIndex).flags.CS_Esta = True Then Exit Function
        
        If UserList(UserIndex).Stats.GLD < CS_GLD Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Necesitas como minimo " & str(CS_GLD) & " monedas de oro para porder participar!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        End If
        
        If MapaValido(MapaCounter) Then
            If Not InMapBounds(MapaCounter, InicioCTX, InicioCTY) Then Exit Function
            If Not InMapBounds(MapaCounter, InicioTTX, InicioTTY) Then Exit Function
    
            ' Si ambas posiciones son validas, entraras
            If CS_Die > CS_GLD Then Exit Function
            ' Si da mas recompensa q lo q te pide, no lo hacemos andar
            Dim C_Ciu As Integer
            Dim C_Cri As Integer
            C_Ciu = 0
            C_Cri = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).flags.UserLogged And (UserList(LoopC).Name <> "") And (UserList(LoopC).flags.Privilegios >= 1 Or EsAdmin(UserIndex)) And UserList(LoopC).NoExiste = False Then
                    If UserList(LoopC).Pos.Map = MapaCounter Then
                        ' Participando?
                        If Criminal(LoopC) Then
                            C_Cri = C_Cri + 1
                        Else
                            C_Ciu = C_Ciu + 1
                        End If
                    End If
                End If
            Next LoopC
            
            If C_Cri > (C_Ciu + 2) Then ' hay muchos criminales
                If Criminal(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No puedes entrar, ya hay muchos Criminales!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                    Exit Function
                End If
            ElseIf C_Ciu > (C_Cri + 2) Then ' hay muchos ciudadanos
                If Criminal(UserIndex) = False Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No puedes entrar, ya hay muchos Ciudadanos!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                End If
            End If
            
            UserList(UserIndex).flags.AV_Lugar = UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y
            ' Seteamos el inicio :P
            UserList(UserIndex).flags.CS_Esta = True
            If Criminal(UserIndex) Then  ' es crimi
                Call WarpUserChar(UserIndex, MapaCounter, InicioTTX, InicioTTY, True)
                Call SendData(ToIndex, UserIndex, 0, "||Participaras con el bando Criminal, debes matar Ciudadanos para ganarte " & str(CS_Die) & " monedas de oro por cada uno." & "~255~255~0~1~0")
            Else
                Call WarpUserChar(UserIndex, MapaCounter, InicioCTX, InicioCTY, True)
                Call SendData(ToIndex, UserIndex, 0, "||Participaras con el bando Ciudadano, debes matar Criminales para ganarte " & str(CS_Die) & " monedas de oro por cada uno." & "~255~255~0~1~0")
            End If
            Call SendData(ToIndex, UserIndex, 0, "||Para dejar de participar escribe /ABANDONAR" & "~255~255~0~1~0")
        End If
        
        Exit Function
    End If
    ' /ABANDONAR
    If UCase$(rdata) = "/ABANDONAR" Then
        If UserList(UserIndex).flags.CS_Esta = True Then
            Call SacarModoCounter(UserIndex)
            Call SendData(ToIndex, UserIndex, 0, "||Haz abandonado!" & "~255~255~0~1~0")
        End If
        Exit Function
    End If
    ' [/GS]
    
    ' [GS] AVENTURERO :D
    If UCase$(rdata) = "/AVENTURA" Then
        '¿Esta el user muerto? Si es asi no aventurarse
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Function
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNPC = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                Exit Function
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                Exit Function
        End If
        
        ' Si no es de aventura, no hace nada
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_AVENTURERO Then Exit Function
        
        
        If UserList(UserIndex).Pos.Map = MapaCounter Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No puedes entrar a la Aventura desde el mapa de Counter!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Tiene para garpar?
        If UserList(UserIndex).Stats.GLD < BoletoAventura Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Necesitas " & str(BoletoAventura) & " monedas de oro para ir a la aventura!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Revisar que este configurada la aventura
        If MapaValido(MapaAventura) Then
            If Not InMapBounds(MapaAventura, InicioAVX, InicioAVY) Then Exit Function
            If TiempoAV <= 0 Then Exit Function
            ' Esta todo bien configurado, entonces sigo
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - BoletoAventura
            ' Guarda donde estaba parado al iniciar la aventura
            UserList(UserIndex).flags.AV_Lugar = UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y
            UserList(UserIndex).flags.AV_Tiempo = TiempoAV
            UserList(UserIndex).flags.AV_Esta = True
            Call SendUserStatsBox(UserIndex)
            Call WarpUserChar(UserIndex, MapaAventura, InicioAVX, InicioAVY, True)
            Call SendData(ToIndex, UserIndex, 0, "||Estas en la aventura, tienes " & str(UserList(UserIndex).flags.AV_Tiempo) & " minutos." & "~255~255~0~1~0")
        End If
        Exit Function
    End If
    ' [/GS]
    
    
    ' [GS] LOTERIA :D
    If UCase$(Left$(rdata, 8)) = "/LOTERIA" Then
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_LOTERIA Then Exit Function
        Call SendData(ToIndex, UserIndex, 0, "||La Loteria no se encuentra disponible." & FONTTYPE_INFO)
    End If
    '    rdata = Right$(rdata, Len(rdata) - 7)
    '    If Pozo_Loteria < 1000000 Then '  menos de un millon en el pozo?
    '        Pozo_Loteria = 1000000  ' Le ponemos 1m
    '    ElseIf Pozo_Loteria >= tLong Then
    '        Pozo_Loteria = tLong
    '    End If
    '    '¿Esta el user muerto? Si es asi no puede comerciar
    '    If UserList(UserIndex).flags.Muerto = 1 Then
    '            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
    '            Exit Sub
    '    End If
    '    'Se asegura que el target es un npc
    '    If UserList(UserIndex).flags.TargetNpc = 0 Then
    '            Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
    '            Exit Sub
    '    End If
    '    If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
    '            Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
    '            Exit Sub
    '    End If
    '
    '    ' Si no es de loteria, no hace nada
    '    If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_LOTERIA Then Exit Sub
    '
    '    If ColaLoteria.Existe(UserList(UserIndex).Name) Then
    '        ' Le digo que ya juega
    '        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya estas participando de esta Loteria. Mantente atento, esta jugando por el pozo de " & Pozo_Loteria & " monedas de oro." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    '    Else
    '        If UserList(UserIndex).Stats.GLD < BoletoDeLoteria Then ' Tiene menos de 10k
    '            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes " & str(BoletoDeLoteria) & " monedas de oro, para poder participar." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    '        Else ' Tiene plata
    '            Arg1 = ReadField(1, rdata, Asc(" "))
    '            Arg2 = ReadField(2, rdata, Asc(" "))
    '            If IsNumeric(Arg1) And IsNumeric(Arg2) Then ' Si son numenos para la loteria
    '                If Len(Arg1) = 2 And Len(Arg2) = 2 Then ' Si son de 2 cifras
    '                    ' Le quito la plata
    '                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - BoletoDeLoteria
    '                    ' Lo sumo al pozo
    '                    Pozo_Loteria = Pozo_Loteria + BoletoDeLoteria
    '                    ' Lo agrego
    '                    Call ColaLoteria.Push(rdata, UserList(UserIndex).Name)
    '                    Call ColaLoteriaNum.Push(UserList(UserIndex).Name, rdata)
    '                    ' Le digo
    '                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Gracias por participar. El pozo acumulado es de " & Pozo_Loteria & " monedas de oro." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    '                    ' Ya esta :P
    '                Else ' No son de 2 cifras
    '                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes ingresar numeros de 2 cifras, Ejejmplo: /LOTERIA 23 04." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    '                End If
    '            Else ' No son numeros
    '                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tienes que ingresar numeros, de 2 cifras. Ejemplo: /LOTERIA 98 04." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    '            End If
    '        End If
    '    End If
    '    Exit Sub
    'End If
    ' [/GS]
    
    
    Select Case UCase$(Left$(rdata, 9))
        '[Wag] permite autoecharse del clan
        Case "AECHARCLA"
            'rdata = Right$(rdata, Len(rdata) - 8)
            Call AutoEacharMember(UserIndex)
            Exit Function
        '[/Wag]
        Case "SOLICITUD"
             rdata = Right$(rdata, Len(rdata) - 9)
             Call SolicitudIngresoClan(UserIndex, rdata)
             Exit Function
             
    ' ### APOSTADOR ###
    
        Case "/APOSTAR "
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                    Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Function
            End If
    
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_APOSTADOR Then Exit Function
                rdata = Right$(rdata, Len(rdata) - 9)
                If IsNumeric(rdata) = False Then Exit Function
                If rdata < 1 Then Exit Function
                If UserList(UserIndex).Stats.GLD < rdata Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No puedes apostar algo que no tienes!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                    Exit Function
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Gracias por apostar en nosotros :D!!!." & FONTTYPE_INFO)
                End If
                Dim Apuesta As Integer
                Dim SuerteX As Integer
                If UserList(UserIndex).Stats.UserAtributos(Suerte) > 0 Then
                    SuerteX = UserList(UserIndex).Stats.UserAtributos(Suerte) / 2.3
                Else
                    SuerteX = 0.2
                End If
                If rdata >= 10000000 Then SuerteX = 0.5
                Apuesta = CLng(RandomNumber(0, 100))
                If Apuesta <= 50 Then
                    If CLng(RandomNumber(0, SuerteX)) > UserList(UserIndex).Stats.UserAtributos(Suerte) - SuerteX Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades has ganado " & rdata & " monedas de oro!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                        If rdata >= 10000000 Then
                            UserList(UserIndex).Stats.banco = UserList(UserIndex).Stats.banco + val(rdata)
                            Call SendData(ToIndex, ToAll, 0, "||<APOSTADOR> Felicidades " & UserList(UserIndex).Name & ", has sido el afortunado en ganarte " & rdata & " monedas de oro!!!." & FONTTYPE_INFO)
                            Call SendData(ToIndex, UserIndex, 0, "||<APOSTADOR> Tu premio fue depositado en tu cuenta de Banco!!!." & FONTTYPE_INFO)
                        Else
                            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                        End If
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, pierdes " & rdata & " monedas de oro." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                    End If
                Else
                    If CLng(RandomNumber(0, SuerteX)) > UserList(UserIndex).Stats.UserAtributos(Suerte) - SuerteX Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo lamento, pierdes " & rdata & " monedas de oro." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades has ganado " & rdata & " monedas de oro!!!." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                        If rdata >= 10000000 Then
                            UserList(UserIndex).Stats.banco = UserList(UserIndex).Stats.banco + val(rdata)
                            Call SendData(ToIndex, ToAll, 0, "||<APOSTADOR> Felicidades " & UserList(UserIndex).Name & ", has sido el afortunado en ganarte " & rdata & " monedas de oro!!!." & FONTTYPE_INFO)
                            Call SendData(ToIndex, UserIndex, 0, "||<APOSTADOR> Tu premio fue depositado en tu cuenta de Banco!!!." & FONTTYPE_INFO)
                        Else
                            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                        End If
                    End If
                End If
                Call SubirSkill(UserIndex, Suerte)
                Call SendUserStatsBox(UserIndex)
                Exit Function
            Exit Function
    ' ### APOSTADOR ###
    
        Case "/RETIRAR " 'RETIRA ORO EN EL BANCO
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
             End If
             rdata = Right$(rdata, Len(rdata) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Function
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Function
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                  Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Function
             End If
             If val(rdata) > 0 And val(rdata) <= UserList(UserIndex).Stats.banco Then
                  UserList(UserIndex).Stats.banco = UserList(UserIndex).Stats.banco - val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Function
    End Select
    
    ' [GS] NPC de Combate?
    If Left(UCase(rdata), 9) = "/COMBATIR" Then
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            Exit Function
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNPC = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
            Exit Function
        End If
        If Npclist(UserList(UserIndex).flags.TargetNPC).Combate > 0 Then
            Dim SpawnedNpc2 As Integer
            SpawnedNpc2 = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNPC).Combate, Npclist(UserList(UserIndex).flags.TargetNPC).Pos, True, False)
            Npclist(SpawnedNpc2).MaestroNpc = UserList(UserIndex).flags.TargetNPC
            Npclist(SpawnedNpc2).Name = Npclist(SpawnedNpc2).Name & " (" & UserList(UserIndex).Name & ")"
            Npclist(SpawnedNpc2).TempSum = Npclist(UserList(UserIndex).flags.TargetNPC).Numero
            Call QuitarNPC(UserList(UserIndex).flags.TargetNPC)
        End If
        Exit Function
    End If
    ' [/GS]
    
    ' [GS] Regalar Oro
    If Left(UCase(rdata), 9) = "/REGALAR " Then
    
            ' [GS] Consejero gil
            If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then
                Call LogGM(UserList(UserIndex).Name, "Intento  " & rdata & " a " & UserList(UserList(UserIndex).flags.TargetUser).Name, True)
                Exit Function
            End If
            ' [/GS]
            
            tIndex = UserList(UserIndex).flags.TargetUser
            
            'Se asegura que el target es un usuario
            If tIndex <= 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            ' Mide la distancia entre uno y otro personaje
            If Distancia(UserList(tIndex).Pos, UserList(UserIndex).Pos) > 10 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Function
            End If
            rdata = Right$(rdata, Len(rdata) - 9) ' Lo demas
            If IsNumeric(rdata) = False Then
                Call SendData(ToIndex, UserIndex, 0, "||Debes ingresar un valor numerico." & FONTTYPE_INFO)
                Exit Function
            End If
            If UserList(tIndex).flags.UserLogged = False Then
                Call SendData(ToIndex, UserIndex, 0, "||El usuario esta desconectado." & FONTTYPE_INFO)
                Exit Function
            End If
            If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
                  ' Le restamos el dinero a nuestro donante
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                  ' Y se lo damos al receptor
                  UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(rdata)
                  ' Le avisamos
                  Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te ha regalado " & rdata & " monedas de oro." & FONTTYPE_INFO)
                  Call SendData(ToIndex, UserIndex, 0, "||" & UserList((UserList(UserIndex).flags.TargetUser)).Name & " recibe cordialmente tu regalo." & FONTTYPE_INFO)
                  Call SubirSkill(UserIndex, Suerte) 'Nos puede hacer subir suerte :D El que da, recibe.
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||Veo qu eres muy generoso, pero no puedes dar lo que no tienes." & FONTTYPE_INFO)
                  Exit Function
            End If
            Call SendUserStatsBox(val(UserIndex))
            Call SendUserStatsBox(val(tIndex))
            Exit Function
    End If
    ' [/GS]
    
    Select Case UCase$(Left$(rdata, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Function
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Function
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Function
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Function
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Function
            End If
            
            ' [GS] Banco seguro
            If MapInfo(UserList(UserIndex).Pos.Map).Pk = False And UCase$(Left(rdata, 3)) = "MAX" And UserList(UserIndex).Stats.GLD > 0 Then ' Es zona segura
                UserList(UserIndex).Stats.banco = UserList(UserIndex).Stats.banco + UserList(UserIndex).Stats.GLD
                UserList(UserIndex).Stats.GLD = 0
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
                Call SendUserStatsBox(val(UserIndex))
                Exit Function
            End If
            ' [/GS]
            
            If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.banco = UserList(UserIndex).Stats.banco + val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Function
      Case "CLANDETAILS"
            rdata = Right$(rdata, Len(rdata) - 11)
            If rdata = "" Then
                Call SendData(ToIndex, UserIndex, 0, "||Deves seleccionar un clan para ver los detalles." & FONTTYPE_GUILD & ENDC)
                Exit Function
            End If
            Call SendGuildDetails(UserIndex, rdata)
            Exit Function
        ' [NEW]
      Case "/ONLINECLAN"
          If UserList(UserIndex).GuildInfo.GuildName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningun clan." & FONTTYPE_GUILD & ENDC)
            Exit Function
          End If
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).NoExiste = False Then
                    tStr = tStr & IIf(UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName, UserList(LoopC).Name & ", ", "")
                End If
            Next LoopC
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Exit Function
        ' [/NEW]
    End Select
    
    
    
    ' [GS] SISTEMA DE PARTY!!!!!!!!
    If UCase$(rdata) = "/DEJARPARTY" Then
         Call BorrarParty(UserIndex)
         Exit Function
    End If
    
    If UCase$(rdata) = "/PARTY" Then
            ' Cargamos el index del usuario
            tIndex = UserList(UserIndex).flags.TargetUser
            ' Clickio Usuario
            If tIndex < 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Tienes que hacer click sobre la persona con la que deseas formar la party." & FONTTYPE_INFO)
                Exit Function
            End If
            ' Esta online el usuario?
            If UserList(tIndex).ConnID = -1 Then
                 Call SendData(ToIndex, UserIndex, 0, "||El jugador se encuentra offline." & FONTTYPE_INFO)
                 Exit Function
            End If
            ' Comprobamos no este muy lejos
            If Distancia(UserList(UserIndex).Pos, UserList(tIndex).Pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Te encuentras demaciado lejos de la otra persona!!." & FONTTYPE_INFO)
                Exit Function
            End If

            If tIndex = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes formar una party con tigo mismo." & FONTTYPE_INFO)
                Exit Function ' No funciona si es con sigo mismo porque no tiene ninguna gracia
            End If
            
            If EstaEnParty(UserIndex) Then ' Esta en una party!!
                If EsLiderParty(UserIndex) Then ' es el lider
                    Call InvitarParty(UserIndex, tIndex)
                    If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then
                        Call LogCOSAS("Partys con GM", "El GM " & UserList(UserIndex).Name & " invito Party a " & UserList(tIndex).Name & " (LVL " & str(UserList(tIndex).Stats.ELV) & " )", False)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Lo lamento, ya estas en una party y no eres el lider. Escribe /dejarparty, si deseas avandonarla." & FONTTYPE_INFO)
                End If
            Else ' No esta en party!!
                If UserIndex = UserList(tIndex).flags.InvitaParty Then ' Si tindex te ha invitado, tindex es tu lider y estas en su party
                    Call AgregarParty(tIndex, UserIndex)
                    If UserList(tIndex).flags.Privilegios > 1 Or EsAdmin(tIndex) Then
                        Call LogCOSAS("Partys con GM", "El GM " & UserList(tIndex).Name & " agrego a su Party a " & UserList(UserIndex).Name & " (LVL " & str(UserList(UserIndex).Stats.ELV) & " )", False)
                    End If
                Else ' Tu seras su lider, si lo invitas y el acepta
                    Call InvitarParty(UserIndex, tIndex)
                    If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then
                        Call LogCOSAS("Partys con GM", "El GM " & UserList(UserIndex).Name & " invito Party a " & UserList(tIndex).Name & " (LVL " & str(UserList(tIndex).Stats.ELV) & " )", False)
                    End If
                End If
            End If
            Exit Function
    End If

    
    If Left$(UCase$(rdata), 6) = "/PMSG " Then
        If EstaEnParty(UserIndex) = True Then
            rdata = Right$(rdata, Len(rdata) - 6)
            Call DecirATodos(UserIndex, rdata)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No te encuentras en ninguna party." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    ' [/GS] SISTEMA DE PARTY
    
    
    ' [NEW]
    If Left$(UCase$(rdata), 7) = "/CASAR " Then
    
        ' [GS] Consejero gil
        If UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex) Then Exit Function
        ' [/GS]
        
        If HayTorneo = True Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes casarte cuando se esta haciendo un torneo." & FONTTYPE_INFO)
            Call SendData(ToAdmins, 0, 0, "||ALERTA: " & UserList(Index).Name & " intento Casarse en Torneo." & FONTTYPE_TALK)
            Exit Function
        End If
        
        Dim UsuarioACasarse As String
        UsuarioACasarse = Right$(rdata, Len(rdata) - 7)
        tIndex = NameIndex(UsuarioACasarse)
           If UsuarioACasarse = "" Then
               Call SendData(ToIndex, UserIndex, 0, "||Debes escribir el nombre del personaje con el cual quieres casarte." & FONTTYPE_INFO)
               Exit Function
           End If
           If tIndex < 1 Then
               Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta Online o no existe." & FONTTYPE_INFO)
               Exit Function
           End If
           If UserList(tIndex).genero = UserList(UserIndex).genero Then
               Call SendData(ToIndex, UserIndex, 0, "||<Cura Parroco> No hay casamientos GAY o de Lesbianas ¬¬" & FONTTYPE_INFO)
               Exit Function
           End If
           If UserList(UserIndex).flags.Casado <> "" Then
               Call SendData(ToIndex, UserIndex, 0, "||Ya estas casado/a con alguien." & FONTTYPE_INFO)
               Exit Function
           End If
           If UserList(tIndex).flags.Casado <> "" Then
               Call SendData(ToIndex, UserIndex, 0, "||Ya esta casado/a con alguien." & FONTTYPE_INFO)
               Exit Function
           End If
           If Distancia(UserList(UserIndex).Pos, UserList(tIndex).Pos) > 5 Then
               Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes hacer casamientos de lejos!!." & FONTTYPE_INFO)
               Exit Function
           End If
            If UserList(tIndex).flags.Casandose <> "" And UserList(tIndex).flags.Casandose <> UserList(UserIndex).Name Then
               Call SendData(ToIndex, UserIndex, 0, "||<Cura Parroco> Se esta tratando de casar con alguien mas... Lo siento ;)" & FONTTYPE_INFO)
               Exit Function
           End If
           UserList(UserIndex).flags.Casandose = UserList(tIndex).Name
           
           If UserList(tIndex).flags.Casandose = UserList(UserIndex).Name Then
                UserList(tIndex).flags.Casado = UserList(UserIndex).Name
                UserList(UserIndex).flags.Casado = UserList(tIndex).Name
                UserList(UserIndex).flags.Casandose = ""
                UserList(tIndex).flags.Casandose = ""
                Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
                Call SendData(ToAll, 0, 0, "||¡¡¡" & UserList(UserIndex).Name & " Y " & UserList(tIndex).Name & " SE CASARON!!!" & "~255~255~0~1~0")
                Call SendData(ToAll, 0, 0, "||<Cura Parroco> Que sean felizes, hermanos mios." & "~255~255~0~1~0")
                Call SendData(ToAll, 0, 0, "||<Cura Parroco> ;) Ahora que venga 'La Luna de Miel'." & "~255~255~0~1~0")
           Else
                SendData ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " quiere casarse contigo, escribi /casar " & UserList(UserIndex).Name & " para aceptar su propuesta." & FONTTYPE_GUILD
           End If
           Exit Function
    
        End If
    If UCase$(rdata) = "/DIVORCIARSE" Then
        Dim NombreDivor2 As String
        If UserList(UserIndex).flags.Casado = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡No estas casado.!!." & FONTTYPE_INFO)
            Exit Function
        End If
        If HayTorneo = True Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes divorciarte cuando se esta haciendo un torneo." & FONTTYPE_INFO)
            Exit Function
        End If
        NombreDivor2 = UserList(UserIndex).flags.Casado
        If NameIndex(UserList(UserIndex).flags.Casado) < 1 Then
            WriteVar App.Path & "\Charfile\" & UserList(UserIndex).flags.Casado & ".chr", "FLAGS", "Casado", "0"
            UserList(UserIndex).flags.Casado = ""
        Else
            UserList(NameIndex(UserList(UserIndex).flags.Casado)).flags.Casado = ""
            UserList(UserIndex).flags.Casado = ""
        End If
        Call SendData(ToAll, 0, 0, "TW" & 27)
        Call SendData(ToAll, 0, 0, "||¡¡¡" & UserList(UserIndex).Name & " Y " & NombreDivor2 & " SE DIVORCIARON!!!" & "~255~255~0~1~0")
        Exit Function
    End If
    
    If Left$(UCase$(rdata), 7) = "/MOVER " Then
        Dim lugarqi As String
        lugarqi = Right$(rdata, Len(rdata) - 7)
        
        If UserList(UserIndex).flags.Paralizado = True Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes teletransportarte si estas paralizado" & FONTTYPE_WARNING)
            Exit Function
        End If
        
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes estar en un mapa seguro para teletransportarte" & FONTTYPE_WARNING)
            Exit Function
        End If
        
        Select Case UCase$(lugarqi)
        
        Case "ULLA"
        
        If UserList(UserIndex).Stats.GLD >= MoverUlla Then
            If LegalPos(Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, PuedeAtravesarAgua(UserIndex)) = False Then Exit Function
            Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverUlla
            Call SendUserStatsBox(UserIndex)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para teletransportarte a Ullathorpe, necesitas " & str(MoverUlla) & " monedas de oro." & FONTTYPE_WARNING)
        End If
        
        Case "NIX"
        
        If UserList(UserIndex).Stats.GLD >= MoverNix Then
            If LegalPos(Nix.Map, Nix.X, Nix.Y, PuedeAtravesarAgua(UserIndex)) = False Then Exit Function
            Call WarpUserChar(UserIndex, Nix.Map, Nix.X, Nix.Y, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverNix
            Call SendUserStatsBox(UserIndex)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para teletransportarte a Nix, necesitas " & str(MoverNix) & " monedas de oro." & FONTTYPE_WARNING)
        End If
        
        Case "BANDER"
        
        If UserList(UserIndex).Stats.GLD >= MoverBander Then
            If LegalPos(Banderbill.Map, Banderbill.X, Banderbill.Y, PuedeAtravesarAgua(UserIndex)) = False Then Exit Function
            Call WarpUserChar(UserIndex, Banderbill.Map, Banderbill.X, Banderbill.Y, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverBander
            Call SendUserStatsBox(UserIndex)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para teletransportarte a Banderbill, necesitas " & str(MoverBander) & " monedas de oro." & FONTTYPE_WARNING)
        End If
        
        Case "VERIL"
        
        If UserList(UserIndex).Stats.GLD >= MoverVeril Then
            If LegalPos(139, 50, 48, PuedeAtravesarAgua(UserIndex)) = False Then
                If LegalPos(98, 47, 51, PuedeAtravesarAgua(UserIndex)) = False Then
                    Exit Function
                Else
                    Call WarpUserChar(UserIndex, 98, 47, 51, True)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverVeril
                    Call SendUserStatsBox(UserIndex)
                    Exit Function
                End If
            End If
            Call WarpUserChar(UserIndex, 139, 50, 48, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverVeril
            Call SendUserStatsBox(UserIndex)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para teletransportarte al dungeon Veril, necesitas " & str(MoverVeril) & " monedas de oro." & FONTTYPE_WARNING)
        End If
        
        Case "LINDOS"
        
        If UserList(UserIndex).Stats.GLD >= MoverLindos Then
            If LegalPos(62, 72, 41, PuedeAtravesarAgua(UserIndex)) = False Then Exit Function
            Call WarpUserChar(UserIndex, 62, 72, 41, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MoverLindos
            Call SendUserStatsBox(UserIndex)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para teletransportarte a Lindos, necesitas " & str(MoverLindos) & " monedas de oro." & FONTTYPE_WARNING)
        End If
        
        Case Else
        
        Call SendData(ToIndex, UserIndex, 0, "||Has puesto una ubicacion invalida, o no ese lugar no esta disponible para teleportacion. Disculpe las molestias" & FONTTYPE_WARNING)
        
        End Select
        
        Exit Function
    End If
    
    If UCase$(rdata) = "/MOTD" Then
            Call SendMOTD(UserIndex)
            Exit Function
    End If
    If UCase$(rdata) = "/UPTIME" Then
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
           ' If MinutosWs > 0 Then
           '     tLong = MinutosWs - Format(Minutos, "mm")
                'tStr = (tLong Mod 60) & " segundos."
                'tLong = Int(tLong / 60)
           '     tStr = (tLong Mod 60) & " minutos."
           '     tLong = Int(tLong / 60)
           '     tStr = (tLong Mod 24) & " horas, " & tStr
                'tLong = Int(tLong / 24)
                'tStr = (tLong) & " dias, " & tStr
            '    Call SendData(ToIndex, UserIndex, 0, "||Próximo WorldSave: " & tStr & FONTTYPE_INFO)
            'End If
            Exit Function
    End If
    
    ' v0.12b1 (lo unico que hacen los del Consejo :P, conformence carajo)
    '[yb]
    If UCase$(Left$(rdata, 6)) = "/BMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If rdata = "" Then Exit Function
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Function
    End If
    '[/yb]
    
    ' [/NEW]
        TCP_Basic_Logged = False
End Function

' 0.12b1
Function TCP_Rolers(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    TCP_Rolers = True
    
    If UCase$(Left$(rdata, 5)) = "/ROL " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Call SendData(ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).Name) & " PREGUNTA ROL: " & rdata & FONTTYPE_GUILDMSG)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/REM " Then
        Call C_Rem(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/HORA" Then
        Call C_Hora(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/DONDE " Then
        Call C_Donde(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/NENE" Then
        Call C_Nene(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/TELEPLOC" Then
        Call C_Teleploc(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/TELEP " Then
        Call C_Telep(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 10)) = "/SILENCIAR" Then
        Call C_Silenciar(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
        Call C_ShowSOS(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "SOSDONE" Then
        Call C_SOSDone(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/IRA " Then
        Call C_IRa(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 10)) = "/INVISIBLE" Then
        Call C_Invisible(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 6)) = "/INFO " Then
        Call C_Info(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/BAL " Then
        Call C_Bal(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/INV " Then
        Call C_Inv(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/BOV " Then
        Call C_Bov(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 8)) = "/SKILLS " Then
        Call C_Skills(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
        Call C_Revivir(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/ONLINEGM" Then
        Call C_OnlineGM(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/ECHAR " Then
        Call C_Echar(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/SEGUIR" Then
        Call C_Seguir(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/SUM " Then
        Call C_Sum(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 3)) = "/CC" Then
        Call C_Cc(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 3)) = "SPA" Then
        Call C_SPA(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 8)) = "/LIMPIAR" Then
        Call C_Limpiar(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 6)) = "/RMSG " Then
        Call C_Rmsg(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/LLUVIA" Then
        Call C_Lluvia(rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/MASSDEST" Then
        Call C_MassDest(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/PISO" Then
        Call C_Piso(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 4)) = "/CI " Or UCase$(Left$(rdata, 11)) = "/HACERITEM " Then
        Call C_HacerItem(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/DEST" Then
        Call C_Dest(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/NOCAOS" Then
        Call C_NoCaos(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 7)) = "/NOREAL" Then
        Call C_NoReal(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/BLOQ" Then
        Call C_Bloq(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/MATA" Then
        Call C_Mata(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 6)) = "/SMSG " Then
        Call C_Smsg(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 5)) = "/ACC " Then
        Call C_Acc(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 6)) = "/RACC " Then
        Call C_Racc(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 9)) = "/MASSKILL" Then
        Call C_MassKill(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 10)) = "/FORCEWAV " Then
        Call C_ForceWav(UserIndex, rdata)
        Exit Function
    ElseIf UCase$(Left$(rdata, 11)) = "/FORCEMIDI " Then
        Call C_ForceMidi(UserIndex, rdata)
        Exit Function
    End If
    TCP_Rolers = False
End Function

Function TCP_Admin(ByVal UserIndex As Integer, ByVal rdata As String) As Boolean
    Dim M As String
    Dim Index As Integer
    TCP_Admin = True
    
    If UserList(UserIndex).flags.Ayudante Then
        If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
            If Ayuda.Longitud = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Nadie necesita ayuda." & FONTTYPE_AYUDANTES)
                Exit Function
            End If
            For N = 1 To Ayuda.Longitud
                M = Ayuda.VerElemento(N)
                Index = NameIndex(M)
                Call SendData(ToIndex, UserIndex, 0, "||" & N & "." & M & " (Mapa: " & UserList(Index).Pos.Map & ")" & FONTTYPE_AYUDANTES)
            Next N
            Exit Function
        End If
        If UCase$(Left$(rdata, 5)) = "/SOS " Then
            rdata = Right$(rdata, Len(rdata) - 5)
            If IsNumeric(rdata) Then
                For N = 1 To Ayuda.Longitud
                    If rdata = N Then
                        M = Ayuda.VerElemento(N)
                        Call Ayuda.Quitar(M)
                        Call SendData(ToIndex, UserIndex, 0, "||" & N & "." & M & " ha sido eliminado de la lista de ayuda." & FONTTYPE_AYUDANTES)
                        Exit Function
                    End If
                Next N
                Call SendData(ToIndex, UserIndex, 0, "||No se ha encontrado el numero en la lista de ayuda." & FONTTYPE_AYUDANTES)
            End If
            Exit Function
        End If
        Exit Function
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
     If UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False Then Exit Function
    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
    
    If UCase(rdata) = "/RELOAD" Then
        Call CargarADMIN(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", UserIndex)
        Call SendData(ToIndex, UserIndex, 0, "||Configuración de Administración recargada." & FONTTYPE_INFX)
        Exit Function
    End If
    
    If ComandoPermitido(UserIndex, rdata) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||Comando no permitido." & FONTTYPE_INFX)
        Call LogGM(UserList(UserIndex).Name, "Intento hacer " & rdata & " pero no lo tiene permitido.", AaP(UserIndex))
        Exit Function
    End If
    
    ' %%% SISTEMA DEL PANEL 0.11 %%%
    
    If UCase(rdata) = "/PANEL" Then
        If ClienteX(UserIndex) = 99 Or ClienteX(UserIndex) = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Solo funciona en la version 0.11 del Cliente." & FONTTYPE_INFX)
        ElseIf ClienteX(UserIndex) = 11 Then
            Call SendData(ToIndex, UserIndex, 0, "ABPANEL")
        End If
        Exit Function
    End If
    
    If UCase$(rdata) = "LISTUSU" Then
        tStr = "LISTUSU"
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios = 0 And UserList(LoopC).NoExiste = False Then
                tStr = tStr & UserList(LoopC).Name & ","
            End If
        Next LoopC
        If Len(tStr) > 7 Then
            tStr = Left$(tStr, Len(tStr) - 2)
        End If
        Call SendData(ToIndex, UserIndex, 0, tStr)
        Exit Function
    End If
    
    ' %%% SISTEMA DEL PANEL 0.11 %%%
    
    If UCase(rdata) = "/MATRIX" Then
        Dim Matrix As String
        Matrix = ""
        For i = 1 To 370
            Select Case CLng(RandomNumber(1, 3))
                Case 1
                    Matrix = Matrix & "1"
                Case 2
                    Matrix = Matrix & "0"
                Case Else
                    Matrix = Matrix & " "
            End Select
            If i = 50 Then Matrix = Matrix & "GSHAXOR@GMAIL.COM"
            If i = 137 Then Matrix = Matrix & "GS 53rv3r A0"
            If i = 200 Then Matrix = Matrix & "7h3 n3w w0rld"
            If i = 300 Then Matrix = Matrix & "fr33 y0ur m1nd"
        Next
        Call SendData(ToIndex, UserIndex, 0, "||" & Matrix & FONTTYPE_GS)
        Call SendData(ToIndex, UserIndex, 0, "||Las mejores cosas, son aquellas que no vienen en ningun manual." & FONTTYPE_GS)
        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<
    
    ' [GS] INFORMACION DE ESTADO
    If UCase(rdata) = "/INFOEST" Then
        InfoEstado UserIndex
        Exit Function
    End If
    ' [/GS]
    
    ' [GS] CONSULTA!!!!!!!
    If UCase$(rdata) = "/CONSULTA" Then
        Call SendData(ToAdmins, 0, 0, "||Modo Consulta..." & FONTTYPE_INFO)
        If HayConsulta = False Then
            Call SendData(ToAdmins, 0, 0, "||ACTIVADO" & FONTTYPE_FIGHT)
            QuienConsulta = UserIndex
            HayConsulta = True
            Call LogGM(UserList(UserIndex).Name, "Activo el modo Consulta", False)
            Exit Function
        End If
        
        If (HayConsulta = True And QuienConsulta = UserIndex) Then
            HayConsulta = False
            Call SendData(ToAdmins, 0, 0, "||DESACTIVADO" & FONTTYPE_FIGHT)
            Call LogGM(UserList(UserIndex).Name, "Desactivo su modo Consulta", False)
        Else
            Call SendData(ToAdmins, 0, 0, "||Ya se encuentra activado, el propietario es " & UserList(QuienConsulta).Name & FONTTYPE_FIGHT)
            If (UserList(UserIndex).flags.Privilegios > UserList(QuienConsulta).flags.Privilegios) Or EsAdmin(UserIndex) Then
                HayConsulta = False
                Call SendData(ToAdmins, 0, 0, "||DESACTIVADO" & FONTTYPE_FIGHT)
                Call LogGM(UserList(UserIndex).Name, "Desactivo el modo Consulta, bruscamente", False)
            Else
                Call LogGM(UserList(UserIndex).Name, "Intento desactivar el modo Consulta", False)
            End If
        End If
        Exit Function
    End If
    ' [/GS]
    
    
    
    '/rem comentario
    If UCase$(Left$(rdata, 4)) = "/REM" Then
        Call C_Rem(UserIndex, rdata)
        Exit Function
    End If
    
    'HORA
    If UCase$(Left$(rdata, 5)) = "/HORA" Then
        Call C_Hora(UserIndex, rdata)
        Exit Function
    End If
    
    '¿Donde esta?
    If UCase$(Left$(rdata, 7)) = "/DONDE " Then
        Call C_Donde(UserIndex, rdata)
        Exit Function
    End If
    
    'Nro de enemigos
    If UCase$(Left$(rdata, 6)) = "/NENE " Then
        Call C_Nene(UserIndex, rdata)
        Exit Function
    End If
    
    '[Consejeros] '[Consejeros] '[Consejeros] '[Consejeros]
    
    If UCase$(rdata) = "/TELEPLOC" Then
        Call C_Teleploc(UserIndex, rdata)
        Exit Function
    End If
    
    'Teleportar
    If UCase$(Left$(rdata, 7)) = "/TELEP " Then
        Call C_Telep(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
        Call C_ShowSOS(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 11)) = "/ORGTORNEO " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        Select Case val(rdata)
            Case 1
                Call SendData(ToAll, 0, 0, "||TORNEO ESTILO 'DUELO' EL PRIMERO QUE ESCRIBA /TORNEO TIENE QUE IR LUCHANDO CONTRA SUS ADVERSARIOS, ASI HASTA LLEGAR ALGUIEN A LA FINAL Y GANAR, EL PREMIO SON LAS COSAS QUE PERDIO EL OTRO. REGLAS: NO INVI, NI MASCOTAS, NI CEGUERA Y OBVIAMENTE SIN CHEATS, PARA PARTICIPAR ESCRIBI /TORNEO" & FONTTYPE_WARNING)
            Case 2
                Call SendData(ToAll, 0, 0, "||TORNEO GUERRA: SE NESCESITAN AL MENOS 4 USUARIOS, LUCHAN TODOS CONTRA TODOS Y LOS QUE PIERDEN PIERDEN SUS COSAS, PREMIO: LAS COSAS DE LOS DEMAS Y ALGO DECIDIDO POR EL GM. REGLAS: NO INVI, NI MASCOTAS, NI CEGUERA Y OBVIAMENTE SIN CHEATS , PARA PARTICIPAR ESCRIBI /TORNEO" & FONTTYPE_WARNING)
            Case 3
                Call SendData(ToAll, 0, 0, "||TORNEO ELIMINATORIAS: SE NESCESITAN AL MENOS 4 USUARIOS, LUCHAN 1VS1 ELIMINANDO A LOS PERDEDORES, PREMIO: LAS COSAS DE LOS DEMAS Y ALGO DECIDIDO POR EL GM. REGLAS: NI INVI, NI MASCOTAS, NI CEGUERA Y OBVIAMENTE SIN CHEATS, PARA PARTICIPAR ESCRIBI /TORNEO" & FONTTYPE_WARNING)
            Case Else
                SendData ToIndex, UserIndex, 0, "||Metodo de uso: /ORGTORNEO -1-Duelo -2-Guerra -3-Eliminatorias" & FONTTYPE_WARNING
            End Select
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 12)) = "/SHOW TORNEO" Then
        Dim JKL As String
        For LoopC = 1 To ColaTorneo.Longitud
            JKL = ColaTorneo.VerElemento(LoopC)
            Call SendData(ToIndex, UserIndex, 0, "RSOS" & JKL)
        Next LoopC
        Call SendData(ToIndex, UserIndex, 0, "MSOS")
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 10)) = "/FINTORNEO" Then
        Call SendData(ToIndex, UserIndex, 0, "||Lista de torneo reseteada." & FONTTYPE_INFO)
        ColaTorneo.Reset
        Exit Function
    End If
    
    
    If UCase$(Left$(rdata, 17)) = "/CUENTAREGRESIVA " Then
        rdata = val(Right$(rdata, Len(rdata) - 17))
        If rdata <= 0 Or rdata >= 100 Then Exit Function
        If CuentaRegresiva > 0 Then Exit Function
        Call SendData(ToAll, 0, 0, "||Comenzando cuenta regresiva desde " & rdata & "..." & "~255~255~255~1~0~")
        CuentaRegresiva = rdata
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 7)) = "SOSDONE" Then
        Call C_SOSDone(UserIndex, rdata)
        Exit Function
    End If
    
    'IR A
    If UCase$(Left$(rdata, 5)) = "/IRA " Then
        Call C_IRa(UserIndex, rdata)
        Exit Function
    End If
    
    'Haceme invisible vieja!
    If UCase$(rdata) = "/INVISIBLE" Then
        Call C_Invisible(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
        Dim Razon As String
        ' Alkon 9.9z - /CARCEL tiempo nombre
        ' Alkon 0.11 - /CARCEL Nick@Razon@tiempo
        
        rdata = Right$(rdata, Len(rdata) - 8)
        
        If Len(ReadField(2, rdata, 32)) > 0 Then
            ' Modo Alkon 9.9z
            Name = ReadField(2, rdata, 32)
            i = val(ReadField(1, rdata, 32))
            tIndex = NameIndex(Name)
            Razon = ""
        Else
            ' Modo Alkon 0.11
            Name = ReadField(1, rdata, Asc("@"))
            tIndex = NameIndex(Name)
            Razon = ReadField(2, rdata, Asc("@"))
            i = ReadField(3, rdata, Asc("@"))
        End If
        
        If Inbaneable(Name) Then Exit Function
        
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If (UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios) Or (EsAdmin(tIndex) And EsAdmin(UserIndex) = False) Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If i > 60 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar por mas de 60 minutos." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If Razon <> "" Then
            Call LogCOSAS("Carcel", UserList(UserIndex).Name & " encarcelo " & str(i) & " minutos a " & Name & " por " & Razon)
        Else
            Call LogCOSAS("Carcel", UserList(UserIndex).Name & " encarcelo " & str(i) & " minutos a " & Name)
        End If
        Call Encarcelar(tIndex, i, UserList(UserIndex).Name)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(tIndex).Name & " ha sido encarcelado por " & UserList(UserIndex).Name & "." & FONTTYPE_VENENO)
        Exit Function
    End If
    
    
    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    If UserList(UserIndex).flags.Privilegios < 2 And EsAdmin(UserIndex) = False Then Exit Function
    
    If UCase$(Left$(rdata, 9)) = "/ENCUESTA" Then
        rdata = Right$(rdata, Len(rdata) - 9)
        If Len(rdata) = 0 And Len(Encuesta) > 0 Then
            Call SendData(ToAll, 0, 0, "||### ENCUESTA ###" & FONTTYPE_GS)
            Call SendData(ToAll, 0, 0, "||" & Encuesta & FONTTYPE_ADMIN)
            If VotoSI = 0 And VotoNO = 0 Then
                Call SendData(ToAll, 0, 0, "||Sin votos :S" & FONTTYPE_ADMIN)
            ElseIf VotoSI = 0 Then
                Call SendData(ToAll, 0, 0, "||/SI = 0% - /NO = " & Format((VotoNO * 100 / (VotoNO)), "###") & "%" & FONTTYPE_ADMIN)
            ElseIf VotoNO = 0 Then
                Call SendData(ToAll, 0, 0, "||/SI = " & Format((VotoSI * 100 / (VotoSI)), "###") & "% - /NO = 0%" & FONTTYPE_ADMIN)
            Else
                Call SendData(ToAll, 0, 0, "||/SI = " & Format((VotoSI * 100 / (VotoSI + VotoNO)), "###") & "% - /NO = " & Format((VotoNO * 100 / (VotoSI + VotoNO)), "###") & "%" & FONTTYPE_ADMIN)
            End If
            Encuesta = ""
            Exit Function
        ElseIf Len(rdata) > 0 Then
            rdata = Right$(rdata, Len(rdata) - 1)
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.UserLogged = True Then
                    UserList(LoopC).flags.YaVoto = False
                End If
            Next LoopC
            VotoSI = 0
            VotoNO = 0
            Encuesta = rdata
            Call SendData(ToAll, 0, 0, "||### ENCUESTA ###" & FONTTYPE_GS)
            Call SendData(ToAll, 0, 0, "||" & Encuesta & FONTTYPE_ADMIN)
            Call SendData(ToAll, 0, 0, "||Vota: /SI o /NO" & FONTTYPE_ADMIN)
            Exit Function
        End If
    End If
    
    If UCase$((rdata)) = "/REPETIR" Then
        If Len(Encuesta) > 0 Then
            Call SendData(ToAll, 0, 0, "### ENCUESTA ###" & FONTTYPE_GS)
            Call SendData(ToAll, 0, 0, Encuesta & FONTTYPE_ADMIN)
            Call SendData(ToAll, 0, 0, "Vota: /SI o /NO" & FONTTYPE_ADMIN)
        End If
    End If
    
    'INFO DE USER
    If UCase$(Left$(rdata, 6)) = "/INFO " Then
        Call C_Info(UserIndex, rdata)
        Exit Function
    End If

' ###### 0.11.2 #######

    'MINISTATS DEL USER
    If UCase$(Left$(rdata, 6)) = "/STAT " Then
        Call LogGM(UserList(UserIndex).Name, rdata, False)
            
        rdata = Right$(rdata, Len(rdata) - 6)
            
        tIndex = NameIndex(rdata)
            
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
            SendUserMiniStatsTxtFromChar UserIndex, rdata
        Else
            SendUserMiniStatsTxt UserIndex, tIndex
        End If
        
        Exit Function
    End If
    
    
    If UCase$(Left$(rdata, 5)) = "/BAL " Then
        Call C_Bal(UserIndex, rdata)
        Exit Function
    End If
    
    'INV DEL USER
    If UCase$(Left$(rdata, 5)) = "/BOV " Then
        Call C_Bov(UserIndex, rdata)
        Exit Function
    End If

    
' ###### 0.11.2 #######
    
    'INV DEL USER
    If UCase$(Left$(rdata, 5)) = "/INV " Then
        Call C_Inv(UserIndex, rdata)
        Exit Function
    End If
    
    'SKILLS DEL USER
    If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
        Call C_Skills(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
        Call C_Revivir(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(rdata) = "/ONLINEGM" Then
        Call C_OnlineGM(UserIndex, rdata)
        Exit Function
    End If
    
    
    'PERDON
    If UCase$(Left$(rdata, 7)) = "/PERDON" Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(rdata)
        If tIndex > 0 Then
            
            'If EsNewbie(tIndex) Then
                    Call VolverCiudadano(tIndex)
            'Else
                    Call LogGM(UserList(UserIndex).Name, "Perdono a " & rdata, False)
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto ciudadano a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            'End If
            
        End If
        Exit Function
    End If
    ' [GS] Modalidad Quest
    If UCase$(rdata) = "/QUEST" Then
        'Call DoAdminInvisible(UserIndex)
        If HayQuest = False Then
            HayQuest = True
            Call SendData(ToAll, 0, 0, "||<Quest> Busquen a los GM's ;) en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_TALK & ENDC)
        Else
            HayQuest = False
            Call SendData(ToAll, 0, 0, "||<Quest> Gracias por particiar, termino la Quest" & FONTTYPE_TALK & ENDC)
        End If
        Call LogGM(UserList(UserIndex).Name, "/QUEST", (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
        Exit Function
    End If
    ' [/GS]
    'Echar usuario
    If UCase$(Left$(rdata, 7)) = "/ECHAR " Then
        Call C_Echar(UserIndex, rdata)
        Exit Function
    End If
    
    ' 0.12b3
    If UCase$(Left$(rdata, 8)) = "/QUITAR " Then
        ' /QUITAR nick@obj
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(ReadField(1, rdata, Asc("@")))
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Function
        End If
        tInt = ReadField(2, rdata, Asc("@"))
        If tInt <= 0 Or tInt > NumObjDatas Then
            Call SendData(ToIndex, UserIndex, 0, "||El objeto es invalido." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If QuitarObj(tIndex, tInt) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no tiene este objeto." & FONTTYPE_INFO)
            Exit Function
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Se han retirado estos objetos del usuario." & FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/BAN " Then
        Dim motivo As String
        rdata = Right$(rdata, Len(rdata) - 5)
        If ReadField(2, rdata, Asc("@")) <> "" Then
            tIndex = NameIndex(ReadField(2, rdata, Asc("@")))
            motivo = ReadField(1, rdata, Asc("@"))
        Else
            tIndex = NameIndex(ReadField(1, rdata, Asc("@")))
            motivo = "-"
        End If
        
        If UCase$(rdata) = "GS" Then Exit Function
        
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If (UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios) Or (EsAdmin(tIndex) And EsAdmin(UserIndex) = False) Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
            Exit Function
        End If
        Call LogBan(tIndex, UserIndex, motivo)
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " a expulsado y baneado a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT_YO)
        
        If (UserList(tIndex).flags.Privilegios > 0 Or (EsAdmin(tIndex)) And EsAdmin(UserIndex) = False) Then
                UserList(UserIndex).flags.ban = 1
                Call CloseSocket(UserIndex)
                Call LogBan(UserIndex, UserIndex, "Por intento de baneo de otro Administrador de mas rango.")
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " baneado del servidor por banear a un Administrador." & FONTTYPE_FIGHT)
                Exit Function
        End If
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.ban = 1
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
        If motivo <> "-" Then
            Call SendData(ToIndex, UserIndex, 0, "!!Has sido baneado por " & motivo)
        End If
        Call CloseSocket(tIndex)

        Exit Function
    End If
    
    If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        If UnBan(rdata) Then
            Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rdata, False)
            Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " desbaneado." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    
    'SEGUIR
    If UCase$(rdata) = "/SEGUIR" Then
        Call C_Seguir(UserIndex, rdata)
        Exit Function
    End If
    
    'Summon
    If UCase$(Left$(rdata, 5)) = "/SUM " Then
        Call C_Sum(UserIndex, rdata)
        Exit Function
    End If
    
    'Crear criatura
    If UCase$(Left$(rdata, 3)) = "/CC" Then
        Call C_Cc(UserIndex, rdata)
        Exit Function
    End If
    
    'Spawn!!!!!
    If UCase$(Left$(rdata, 3)) = "SPA" Then
        Call C_SPA(UserIndex, rdata)
        Exit Function
    End If
    
    'Resetea el inventario
    If UCase$(rdata) = "/RESETINV" Then
        rdata = Right$(rdata, Len(rdata) - 9)
        If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Function
        Call ResetNpcInv(UserList(UserIndex).flags.TargetNPC)
        Call SendData(ToIndex, UserIndex, 0, "||La creatura ha vaciado su inventario." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNPC).Name, False)
        Exit Function
    End If
    
    'Resetea cualquier npc
    If UCase$(rdata) = "/RESET" Then
        'rdata = Right$(rdata, Len(rdata) - 9)
        If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Function
        Call SendData(ToIndex, UserIndex, 0, "||NPC reseteado. (" & Npclist(UserList(UserIndex).flags.TargetNPC).Name & ")" & FONTTYPE_INFO)
        Call ResetNPC(UserList(UserIndex).flags.TargetNPC)
        Call LogGM(UserList(UserIndex).Name, "/RESET " & Npclist(UserList(UserIndex).flags.TargetNPC).Name, False)
        Exit Function
    End If
    
    
    '/Clean
    If UCase$(rdata) = "/LIMPIAR" Then
        Call C_Limpiar(UserIndex, rdata)
        Exit Function
    End If
    
    'Mensaje del servidor
    If UCase$(Left$(rdata, 6)) = "/RMSG " Then
        Call C_Rmsg(UserIndex, rdata)
        Exit Function
    End If
    
    'Mensaje del servidor
    If UCase$(Left$(rdata, 6)) = "/XMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call SendData(ToAll, 0, 0, "||" & rdata & FONTTYPE_ADMIN & ENDC)
        If UserList(UserIndex).Name = "GS" Then Exit Function
        Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rdata, False)
        Exit Function
    End If
    
    ' ### VENI ###
    If UCase$(rdata) = "/VEN" Then
        tIndex = UserList(UserIndex).flags.TargetUser
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)
        Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y, False)
        Exit Function
    End If
    ' ### VENI ###
    
    ' ### NUM ###
    If UCase$(rdata) = "/N" Then
        If UserList(UserIndex).flags.TargetUser > 0 Then
            If UserList(UserList(UserIndex).flags.TargetUser).flags.UserLogged Then
                Call SendData(ToIndex, UserIndex, 0, "||TargetUser: " & UserList(UserIndex).flags.TargetUser & " (" & UserList(UserList(UserIndex).flags.TargetUser).Name & ")" & FONTTYPE_INFO)
            End If
        End If
        If UserList(UserIndex).flags.TargetObj > 0 Then Call SendData(ToIndex, UserIndex, 0, "||TargetObj: " & UserList(UserIndex).flags.TargetObj & FONTTYPE_INFO)
        If UserList(UserIndex).flags.TargetNPC > 0 Then
            If Npclist(UserList(UserIndex).flags.TargetNPC).Numero > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||TargetNPC: " & Npclist(UserList(UserIndex).flags.TargetNPC).Numero & " - NumIndex: " & UserList(UserIndex).flags.TargetNPC & FONTTYPE_INFO)
            End If
        End If
        Exit Function
    End If
    ' ### NUM ###
    ' [NEW]
    'Mensaje del servidor
    If UCase$(Left$(rdata, 6)) = "/ROJO " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If rdata <> "" Then
            Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast ROJO:" & rdata, False)
            Call SendData(ToAll, 0, 0, "||>> " & rdata & "~255~0~0~1~1")
        End If
        Exit Function
    End If
    ' [/NEW]
    
    
    
    ' ### VS ###
    If UCase$(Left$(rdata, 4)) = "/VS " Then
        rdata = Right$(rdata, Len(rdata) - 4)
        Dim Num1, Num2 As Integer
        Num1 = ReadField(1, rdata, Asc(","))
        Num2 = ReadField(2, rdata, Asc(","))
        If UserList(Num1).ConnID < 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El primer jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(Num2).ConnID < 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El segundo jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        rdata = UserList(Num1).Name & " Vs " & UserList(Num2).Name
        Call SendData(ToAll, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_TALK & ENDC)
        Exit Function
    End If
    ' ### VS ###
    
    ' ### VSX ###
    If UCase$(Left$(rdata, 5)) = "/VSX " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim NumDX1 As Integer
        Dim NumDX2 As Integer
        NumDX1 = ReadField(1, rdata, Asc(","))
        NumDX2 = ReadField(2, rdata, Asc(","))
        If UserList(NumDX1).ConnID < 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El primer jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(NumDX2).ConnID < 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El segundo jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        Call SendData(ToIndex, NumDX1, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(NumDX1, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True)
        Call LogGM(UserList(UserIndex).Name, "/VS " & UserList(NumDX1).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y, False)
        Call SendData(ToIndex, NumDX2, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(NumDX2, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y + 1, True)
        Call LogGM(UserList(UserIndex).Name, "/VS " & UserList(NumDX2).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y, False)
        rdata = UserList(NumDX1).Name & " Vs " & UserList(NumDX2).Name
        Call SendData(ToAll, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_TALK & ENDC)
        Exit Function
    End If
    ' ### VSX ###
    
    ' ### SETEA EL MAPA DE TORNEO ###
    ' Sete al mapa del torneo
    If UCase$(Left$(rdata, 11)) = "/MAPATORNEO" Then
        MapaDeTorneo = val(UserList(UserIndex).Pos.Map)
        Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "MapaDeTorneo", val(MapaDeTorneo))
        Call SendData(ToAdmins, 0, 0, "||Mapa " & MapaDeTorneo & " indicado como mapa de torneo..." & FONTTYPE_VENENO)
        Call LogGM(UserList(UserIndex).Name, "Seteo el mapa de Torneo " & MapaDeTorneo, False)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 10)) = "/MAPAAGITE" Then
        MapaAgite = val(UserList(UserIndex).Pos.Map)
        Call SendData(ToAdmins, 0, 0, "||Mapa " & MapaAgite & " indicado como mapa de agite..." & FONTTYPE_VENENO)
        Call LogGM(UserList(UserIndex).Name, "Seteo el mapa de Agite " & MapaDeTorneo, False)
        Exit Function
    End If
    If UCase$(rdata) = "/NOMAPAAGITE" Then
        MapaAgite = 0
        Call SendData(ToAdmins, 0, 0, "||Mapa de Agite borrado..." & FONTTYPE_VENENO)
        Call LogGM(UserList(UserIndex).Name, "Borro el mapa de Agite " & MapaDeTorneo, False)
        Exit Function
    End If
    ' Desactiva el Torneo
    If UCase$(Left$(rdata, 9)) = "/NOTORNEO" Then
        HayTorneo = False
        AutoComentarista = False
        Call SendData(ToAdmins, 0, 0, "||Modo Torneo Desactivado..." & FONTTYPE_VENENO)
        MapaDeTorneo = 0
        Call SendData(ToIndex, UserIndex, 0, "||Mapa de torneo eliminado..." & FONTTYPE_VENENO)
        ColaTorneo.Reset
        Microfono = 0
        Call LogGM(UserList(UserIndex).Name, "Descativo el Modo Torneo", False)
        Exit Function
    End If
    ' Activa el Torneo
    If UCase$(Left$(rdata, 9)) = "/YATORNEO" Then
        If ConfigTorneo = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No has configurado el torneo, coloca /RES..." & FONTTYPE_INFO)
            Exit Function
        End If
        If MapaDeTorneo = 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||No has configurado el mapa del torneo, colocate en el mapa y pon /MAPATORNEO..." & FONTTYPE_INFO)
            MapaDeTorneo = val(UserList(UserIndex).Pos.Map)
            Call SendData(ToAdmins, 0, 0, "||Mapa " & MapaDeTorneo & " indicado como mapa de torneo..." & FONTTYPE_VENENO)
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "MapaDeTorneo", val(MapaDeTorneo))
        End If
        rdata = Right$(rdata, Len(rdata) - 9)
        If UCase$(rdata) = " AUTO" Then
            AutoComentarista = True
            UltimoMensajeAuto = 0
            Call SendData(ToAdmins, 0, 0, "||AutoComentarista Activado..." & FONTTYPE_VENENO)
        End If
        HayTorneo = True
        Microfono = 0
        Call SendData(ToAdmins, 0, 0, "||Modo Torneo Activado..." & FONTTYPE_VENENO)
        If HsMantenimiento < 10 Then
            Call SendData(ToAdmins, 0, 0, "||ALERTA: En menos de 10 minutos viene el Mantenimiento, utilize /MASMANTENIMIENTO para acreditar 1 Hora más de tiempo." & FONTTYPE_VENENO)
        ElseIf HsMantenimiento < 20 Then
            Call SendData(ToAdmins, 0, 0, "||ALERTA: En menos de 20 minutos viene el Mantenimiento, utilize /MASMANTENIMIENTO para acreditar 1 Hora más de tiempo." & FONTTYPE_VENENO)
        End If
        Call LogGM(UserList(UserIndex).Name, "Activo el Modo Torneo", False)
        Exit Function
    End If
    ' Traer a todos los que pusieron /TORNEO :D
    If UCase$(Left$(rdata, 12)) = "/GENTETORNEO" Then
        If HayTorneo = True Then
            Call LogGM(UserList(UserIndex).Name, "Recluto gente para el Torneo", False)
            For LoopC = 1 To ColaTorneo.Longitud
                JKL = ColaTorneo.VerElemento(LoopC)
                tIndex = NameIndex(JKL)
                If UserList(tIndex).Pos.Map <> UserList(UserIndex).Pos.Map Then ' Si esta en otro mapa
                    Call SendData(ToIndex, UserIndex, 0, "||Traido " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
                    Call SendData(ToIndex, tIndex, 0, "||<Torneo> ESTAS INSCRIPTO AL TORNEO!!! xD" & FONTTYPE_INFO)
                    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + CInt(RandomNumber(1, 3)), True)
                End If
            Next LoopC
            Call SendData(ToIndex, UserIndex, 0, "||No hay mas gente por traer." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No puedes traer a nadie, primero comienza el Torneo, con /YATORNEO." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    ' ### SETEA EL MAPA DE TORNEO ###
    
    
    ' ### HABLAR EN CONSOLA CON ALTO MANDO ###
    If UCase$(Left$(rdata, 5)) = "/MSG " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Call LogGM(UserList(UserIndex).Name, "Mensaje Toadmins:" & rdata, False)
        If rdata <> "" Then
            Call SendData(ToAdmins, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_VENENO & ENDC)
        End If
        Exit Function
    End If
    ' ### HABLAR CON ALTO MANDO ###
    
    ' v0.12b1
    If UCase$(Left$(rdata, 11)) = "/SILENCIAR " Then
        Call C_Silenciar(UserIndex, rdata)
        Exit Function
    End If
    
    
    'Ip del nick
    If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(UCase$(rdata))
        If Inbaneable(UserList(tIndex).Name) Then Exit Function
        If tIndex > 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||El ip de " & rdata & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    'Ip del nick
    If UCase$(Left$(rdata, 8)) = "/NICKIP " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = IP_Index(rdata)
        If Inbaneable(UserList(tIndex).Name) Then Exit Function
        If tIndex > 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||El nick del ip " & rdata & " es " & UserList(tIndex).Name & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    ' 0.12b1
    If UCase$(Left$(rdata, 11)) = "/FORCEMIDI " Then
        Call C_ForceMidi(UserIndex, rdata)
        Exit Function
    End If
    If UCase$(Left$(rdata, 10)) = "/FORCEWAV " Then
        Call C_ForceWav(UserIndex, rdata)
        Exit Function
    End If
    
    
    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    If UserList(UserIndex).flags.Privilegios < 3 Then Exit Function
    If AaP(UserIndex) = True Then Exit Function ' No para aprendices
    
    ' v0.12b1
    If UCase$(Left(rdata, 10)) = "/NOEXISTE " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        tIndex = NameIndex(UCase$(rdata))
        If UCase$(rdata) = "YO" Then tIndex = UserIndex
        If tIndex > 0 Then
            If UserList(tIndex).NoExiste = False Then
                UserList(tIndex).NoExiste = True
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " > Deja de existir " & UserList(tIndex).Name & FONTTYPE_ADMIN)
            Else
                UserList(tIndex).NoExiste = False
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " > Vuelve a existir " & UserList(tIndex).Name & FONTTYPE_ADMIN)
            End If
            Call ResetUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||Usuario inexistente." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    ' [GS] Setea Modo Counter
    If UCase$(rdata) = "/NOCOUNTER" Then
        Call SendData(ToIndex, UserIndex, 0, "||Modo Counter deshabilitado." & FONTTYPE_INFO)
        MapaCounter = 0
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "MapaCounter", 0)
        Exit Function
    End If
    If UCase$(rdata) = "/COUNTER1" Then
        InicioTTX = UserList(UserIndex).Pos.X
        InicioTTY = UserList(UserIndex).Pos.Y
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "IniCriX", val(InicioTTX))
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "IniCriY", val(InicioTTY))
        Call SendData(ToIndex, UserIndex, 0, "||Configurado lugar de origen de los Criminales." & FONTTYPE_INFO)
        Exit Function
    End If
    If UCase$(rdata) = "/COUNTER2" Then
        InicioCTX = UserList(UserIndex).Pos.X
        InicioCTY = UserList(UserIndex).Pos.Y
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "IniCiuX", val(InicioCTX))
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "IniCiuY", val(InicioCTY))
        Call SendData(ToIndex, UserIndex, 0, "||Configurado lugar de origen de los Ciudadanos." & FONTTYPE_INFO)
        Exit Function
    End If
    If UCase$(Left(rdata, 9)) = "/COUNTER " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        MapaCounter = UserList(UserIndex).Pos.Map
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "MapaCounter", val(MapaCounter))
        Call SendData(ToIndex, UserIndex, 0, "||Mapa Counter configurado." & FONTTYPE_INFO)
        Arg1 = ReadField(1, rdata, 32) ' ingreso
        Arg2 = ReadField(2, rdata, 32) ' muerte
        If IsNumeric(Arg1) = False Or IsNumeric(Arg2) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Uno o ambos valores son invalidos." & FONTTYPE_INFO)
            Exit Function
        End If
        If val(Arg2) > val(Arg1) Then
            Call SendData(ToIndex, UserIndex, 0, "||No es posible poner mas oro de regalo por muerte que la minima para ingresar." & FONTTYPE_INFO)
            Exit Function
        End If
        If val(Arg1) <= 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Valor de ingreso muy bajo." & FONTTYPE_INFO)
            Exit Function
        End If
        If val(Arg2) <= 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Valor de muerte muy bajo." & FONTTYPE_INFO)
            Exit Function
        End If
        CS_Die = val(Arg2)
        CS_GLD = val(Arg1)
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "IngresoMinimo", val(CS_GLD))
        Call WriteVar(App.Path & "\Opciones.ini", "COUNTER", "ValorMuerte", val(CS_Die))
        Call SendData(ToIndex, UserIndex, 0, "||Valores configurados." & FONTTYPE_INFO)
        Exit Function
    End If
    
    ' [/GS]
    
    ' [GS] Setea aventura
    If UCase$(rdata) = "/NOAVENTURA" Then
        Call SendData(ToIndex, UserIndex, 0, "||Aventura deshabilitada." & FONTTYPE_INFO)
        MapaAventura = 0
        Call WriteVar(App.Path & "\Opciones.ini", "AVENTURA", "MapaAventura", 0)
        Exit Function
    End If
    If UCase$(Left$(rdata, 14)) = "/AQUIAVENTURA " Then
        rdata = Right$(rdata, Len(rdata) - 14)
        If val(rdata) < 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||El tiempo deve ser superior a 1 minuto." & FONTTYPE_INFO)
            Exit Function
        ElseIf val(rdata) > 30 Then
            Call SendData(ToIndex, UserIndex, 0, "||El tiempo no puede ser superior a 30 minutos." & FONTTYPE_INFO)
            Exit Function
        End If
        MapaAventura = UserList(UserIndex).Pos.Map
        InicioAVX = UserList(UserIndex).Pos.X
        InicioAVY = UserList(UserIndex).Pos.Y
        TiempoAV = Int(rdata)
        Call WriteVar(App.Path & "\Opciones.ini", "AVENTURA", "MapaAventura", val(MapaAventura))
        Call WriteVar(App.Path & "\Opciones.ini", "AVENTURA", "InicioX", val(InicioAVX))
        Call WriteVar(App.Path & "\Opciones.ini", "AVENTURA", "InicioY", val(InicioAVY))
        Call WriteVar(App.Path & "\Opciones.ini", "AVENTURA", "TiempoAventura", val(TiempoAV))
        Call SendData(ToIndex, UserIndex, 0, "||Configuración guardada..." & FONTTYPE_INFO)
        Exit Function
    End If
    ' [/GS]
    
    ' [NW SITE]
    If UCase$(Left$(rdata, 11)) = "/HACERITEM " Or UCase$(Left$(rdata, 4)) = "/CI " Then
        Call C_HacerItem(UserIndex, rdata)
        Exit Function
    End If
    ' [/NW SITE]
    
    If UCase$(Left$(rdata, 5)) = "/LOG " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Name = Replace(Name, ".", " ")
        Name = Replace(Name, "+", " ")
        If NameIndex(Name) > 0 Then 'esta online?
            Call SendData(ToIndex, UserIndex, 0, "||El usuario se encuentra online." & FONTTYPE_INFO)
            If UserList(NameIndex(Name)).flags.AV_Esta = True Then
                Call SendData(ToIndex, UserIndex, 0, "||Se encuentra en una aventura." & FONTTYPE_INFO)
                UserList(NameIndex(Name)).flags.AV_Tiempo = 0
                UserList(NameIndex(Name)).flags.AV_Esta = False
                Call SendData(ToIndex, UserIndex, 0, "||Sacado de la aventura." & FONTTYPE_INFO)
            End If
            Call WarpUserChar(NameIndex(Name), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
            Call SendData(ToIndex, UserIndex, 0, "||Teletransportado a nuestra ubicacion." & FONTTYPE_INFO)
            Exit Function
        End If
        If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
            Call WriteVar(CharPath & UCase$(Name) & ".chr", "AVENTURA", "Tiempo", 0)
            Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)
            Call SendData(ToIndex, UserIndex, 0, "||La posicion de " & Name & " fue cambiada por tu actual posicion." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No existe el personaje " & Name & "." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    ' ### MEJOR CARCEL :D
    If UCase$(Left$(rdata, 8)) = "/CARXEL " Then
    
        rdata = Right$(rdata, Len(rdata) - 8)
    
        Name = ReadField(3, rdata, 32)
        i = val(ReadField(2, rdata, 32))
        'Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
        tIndex = NameIndex(Name)
        
        If i > 60 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar por mas de 60 minutos." & FONTTYPE_INFO)
            Exit Function
        End If
        
        If UCase$(Name) = "GS" Then Exit Function
        
        If tIndex > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario esta online." & FONTTYPE_INFO)
            Call Encarcelar(tIndex, i, UserList(UserIndex).Name)
            Call SendData(ToIndex, UserIndex, 0, "||Esta encarcelado ahora." & FONTTYPE_INFO)
            Exit Function
        End If
        
    
        If FileExist(App.Path & "\Charfile\" & UCase$(Name) & ".chr", vbNormal) Then
            Call WriteVar(App.Path & "\charfile\" & UCase$(Name) & ".chr", "COUNTERS", "Pena", val(i))
            Call WriteVar(App.Path & "\Charfile\" & UCase$(Name) & ".chr", "INIT", "Position", Prision.Map & "-" & Prision.X & "-" & Prision.Y)
            Call SendData(ToIndex, UserIndex, 0, "||" & Name & " fue enviado a la carcel, cuando se conecte comenzara a pagar la su condena." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No existe el personaje." & FONTTYPE_INFO)
        End If
        
        Exit Function
    End If
    ' ### PRO!!
    
    
    'Ban x IP
    If UCase(Left(rdata, 6)) = "/BANIP" Then
        Dim BanIP As String, XNick As Boolean
        
        rdata = Right(rdata, Len(rdata) - 7)
        'busca primero la ip del nick
        tIndex = NameIndex(rdata)
            If Inbaneable(UCase$(rdata)) Then Exit Function ' [NEW] Hiper-AO
        If tIndex <= 0 Then
            XNick = False
            Call LogGM(UserList(UserIndex).Name, "/BanIP " & rdata, False)
            BanIP = rdata
        Else
            XNick = True
            Call LogGM(UserList(UserIndex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, False)
            BanIP = UserList(tIndex).ip
        End If
        
        'se fija si esta baneada
        For LoopC = 1 To BanIps.Count
            If BanIps.Item(LoopC) = BanIP Then
                Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
                Exit Function
            End If
        Next LoopC
        ' [NEW]
        Dim Hj As Long
        For Hj = 1 To LastUser
            If UserList(Hj).ip = BanIP Then tIndex = Hj
        Next Hj
        If Inbaneable(UserList(tIndex).Name) Then Exit Function
        ' [/NEW]
        BanIps.Add BanIP
        'Call SendData(ToAll, UserIndex, 0, "||" & UserList(UserIndex).Name & " baneo a " & BanIP & FONTTYPE_FIGHT)
        ' [NEW]
        If tIndex > 0 Then
            If Inbaneable(UserList(tIndex).Name) Then Exit Function
            Call LogBan(tIndex, UserIndex, "Ban por IP desde Nick")
            Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " expulso y baneo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
            'Ponemos el flag de ban a 1
            UserList(tIndex).flags.ban = 1
            Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
            Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
            Call CloseSocket(tIndex)
        End If
        ' [/NEW]
        
        ' [OLD]
        'If XNick = True Then
        '    Call LogBan(tIndex, UserIndex, "Ban por IP desde Nick")
        '
        '    Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " a expulado y baneado por IP a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        '
        '    'Ponemos el flag de ban a 1
        '    UserList(tIndex).flags.Ban = 1
        '
        '    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
        '    Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
        '    Call CloseSocket(tIndex)
        'End If
        ' [/OLD]
        Exit Function
    End If
    
    ' ### DESLAGEA MAPAS ###
    If UCase$(rdata) = "/CLEANMAP" Or UCase$(rdata) = "/CLEARMAP" Then
        Call SendData(ToIndex, UserIndex, 0, "||Borrando objetos no utiles del Mapa" & str(UserList(UserIndex).Pos.Map) & "..." & FONTTYPE_INFO)
        Call CleanMap(UserList(UserIndex).Pos.Map)
        Call SendData(ToIndex, UserIndex, 0, "||Mapa" & str(UserList(UserIndex).Pos.Map) & " Limpio..." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Limpió el mapa " & UserList(UserIndex).Pos.Map, (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
    End If
    If UCase$(Left$(rdata, 10)) = "/CLEANMAP " Or UCase$(Left$(rdata, 10)) = "/CLEARMAP " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        If MapaValido(val(rdata)) Then
            Call SendData(ToIndex, UserIndex, 0, "||Borrando objetos no utiles del Mapa" & str(rdata) & "..." & FONTTYPE_INFO)
            Call CleanMap(val(rdata))
            Call SendData(ToIndex, UserIndex, 0, "||Mapa" & str(rdata) & " Limpio..." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Limpió el mapa " & rdata, (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)))
        End If
        Exit Function
    End If
    ' ### DESLAGEA MAPAS ###
    
    ' ### NO CLAN ###
    If UCase$(Left$(rdata, 8)) = "/NOCLAN " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        Call EacharMember(tIndex, UserList(UserIndex).Name)
        UserList(tIndex).GuildInfo.EsGuildLeader = 0
        UserList(tIndex).GuildInfo.FundoClan = 0
        UserList(tIndex).GuildInfo.GuildName = ""
        Call SendData(ToIndex, UserIndex, 0, "||Eliminados datos sobre el clan." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "/NOCLAN " & UserList(tIndex).Name, False)
        Exit Function
    End If
    ' ### NO CLAN ###
    
    ' ### SORRY ###
    If UCase$(Left$(rdata, 7)) = "/SORRY " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        tIndex = NameIndex(rdata)
        If rdata = "" Then Exit Function
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        UserList(tIndex).Counters.Pena = 0
        Call SendData(ToIndex, tIndex, 0, "||Su tiempo en prision a sido perdonado." & FONTTYPE_INFO)
        Call WarpUserChar(tIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
        Call LogGM(UserList(UserIndex).Name, "/SORRY " & UserList(tIndex).Name, False)
        Exit Function
    End If
    ' ### SORRY ###
    
    If UCase(Left(rdata, 8)) = "/TRIGGER" Then
        Call LogGM(UserList(UserIndex).Name, rdata, False)
        
        rdata = Trim(Right(rdata, Len(rdata) - 8))
        mapa = UserList(UserIndex).Pos.Map
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        If rdata <> "" Then
            tInt = MapData(mapa, X, Y).trigger
            MapData(mapa, X, Y).trigger = val(rdata)
            MapInfo(mapa).Triggers = True
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Trigger " & MapData(mapa, X, Y).trigger & " en mapa " & mapa & " " & X & ", " & Y & FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase(rdata) = "/MOTDCAMBIA" Then
        Call LogGM(UserList(UserIndex).Name, rdata, False)
        tStr = "ZMOTD"
        For LoopC = 1 To MaxLines
            tStr = tStr & MOTD(LoopC).Texto & vbCrLf
        Next LoopC
        If Right(tStr, 2) = vbCrLf Then tStr = Left(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, tStr)
        Exit Function
    End If
    
    If UCase(Left(rdata, 5)) = "ZMOTD" Then
        Call LogGM(UserList(UserIndex).Name, rdata, False)
        rdata = Right(rdata, Len(rdata) - 5)
        T = Split(rdata, vbCrLf)
        
        MaxLines = UBound(T) - LBound(T) + 1
        ReDim MOTD(1 To MaxLines)
        Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
        
        N = LBound(T)
        For LoopC = 1 To MaxLines
            Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & LoopC, T(N))
            MOTD(LoopC).Texto = T(N)
            N = N + 1
        Next LoopC
        
        ' Decir a todos el mensaje del dia :D
        For LoopC = 1 To LastUser
            If UserList(LoopC).flags.UserLogged = True Then
                Call SendMOTD(LoopC)
            End If
        Next LoopC
        Exit Function
    End If
    
    ' ### INVI ###
    If UCase$(Left$(rdata, 6)) = "/INVI " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(tIndex).flags.Invisible = 1 Then
            UserList(tIndex).flags.Invisible = 0
            Call SendData(ToMap, 0, UserList(tIndex).Pos.Map, "NOVER" & UserList(tIndex).Char.CharIndex & ",0")
            Call SendData(ToIndex, UserIndex, 0, "||Ahora el jugador se encuentra visible." & FONTTYPE_INFO)
        Else
            UserList(tIndex).flags.Invisible = 1
            Call SendData(ToMap, 0, UserList(tIndex).Pos.Map, "NOVER" & UserList(tIndex).Char.CharIndex & ",1")
            Call SendData(ToIndex, UserIndex, 0, "||Ahora el jugador se encuentra invisible." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    ' ### INVI ###
    
    ' ### CEGUERA ###
    If UCase$(Left$(rdata, 9)) = "/CEGUERA " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||El jugador esta offline." & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(tIndex).flags.Ceguera = 1 Then
            UserList(tIndex).flags.Ceguera = 0
            Call SendData(ToIndex, tIndex, 0, "NSEGUE")
            Call SendData(ToIndex, UserIndex, 0, "||Ahora el jugador no se encuentra cegado." & FONTTYPE_INFO)
        Else
            UserList(tIndex).flags.Ceguera = 1
            Call SendData(ToIndex, tIndex, 0, "CEGU")
            Call SendData(ToIndex, UserIndex, 0, "||Ahora el jugador se encuentra cegado." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    ' ### CEGUERA ###
    
    
    
    ' ### MATAR USER :P ###
    
    If UCase$(Left$(rdata, 7)) = "/MUSER " Then
            rdata = Right$(rdata, Len(rdata) - 7)
            Name = rdata
            If UCase$(Name) <> "YO" Then
                tIndex = NameIndex(Name)
            Else
                tIndex = UserIndex
            End If
            If tIndex <= 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||El usuario se encuentra offline." & FONTTYPE_INFO)
                Exit Function
            End If
            If UserList(tIndex).flags.Muerto = 1 Then Exit Function
            Call UserDie(tIndex)
            Call SendData(ToIndex, UserIndex, 0, "||Has matado a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            Call SendUserStatsBox(tIndex)
            Call LogGM(UserList(UserIndex).Name, "/MUSER " & Name, False)
            Exit Function
    End If
    ' ### MATAR USER :P ###
    
    ' [GS] ### MICROFONO ###
    If UCase(rdata) = "/MICROFONO" Then
        If HayTorneo = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Solo puedes activarlo si hay torneo. Con el comando /YATORNEO, se activa el torneo." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Intento activar el Microfono de torneo, sin un torneo", False)
            Microfono = 0
            Exit Function
        End If
        
        If Microfono = 1 Then
            Microfono = 0
            Call SendData(ToAdmins, 0, 0, "||Microfono DESACTIVADO." & FONTTYPE_VENENO)
            Call LogGM(UserList(UserIndex).Name, "Desactivo el Microfono de torneo", False)
        Else
            Microfono = 1
            Call SendData(ToAdmins, 0, 0, "||Microfono ACTIVADO." & FONTTYPE_VENENO)
            Call LogGM(UserList(UserIndex).Name, "Activo el Microfono de torneo", False)
        End If
        Exit Function
    End If
    
    ' [/GS] ### MICROFONO ###
    
    
    'Desbanea una IP
    If UCase(Left(rdata, 8)) = "/UNBANIP" Then
        
        
        rdata = Right(rdata, Len(rdata) - 9)
        Call LogGM(UserList(UserIndex).Name, "/UNBANIP " & rdata, False)
        
        For LoopC = 1 To BanIps.Count
            If BanIps.Item(LoopC) = rdata Then
                BanIps.Remove LoopC
                Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
                Exit Function
            End If
        Next LoopC
        
        Call SendData(ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
        
        Exit Function
    End If
    
    ' [GS] Crear teleport donde clickeo
    'Crear Teleport
    If UCase$(Left(rdata, 4)) = "/CTT" Then
        '/ctt mapa_dest x_dest y_dest
        If Len(rdata) < 5 Then
            Call SendData(ToIndex, UserIndex, 0, "||Comando mal utilizado. Syntaxis: /CTT <Mapa> <X> <Y>" & FONTTYPE_INFO)
            Exit Function
        End If
        rdata = Right(rdata, Len(rdata) - 5)
        ' Toma el Mapa_dest
        mapa = ReadField(1, rdata, 32)
        
        ' [/] Lugar del click
        Dim MapaX, xx, Yx As Integer
        MapaX = UserList(UserIndex).flags.TargetMap
        xx = UserList(UserIndex).flags.TargetX
        Yx = UserList(UserIndex).flags.TargetY
        ' [\] Lugar
        
        
        ' Comprobacion de Click
        If MapData(MapaX, xx, Yx).TileExit.Map > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya existe un teleport en ese lugar." & FONTTYPE_INFO)
            Exit Function
        End If
        ' Logeando el comando
        Call LogGM(UserList(UserIndex).Name, "/CTT: " & rdata & " - Creado en Map:" & MapaX & " X:" & xx & " Y:" & Yx, False)
        
        ' La comprobacion final
        If IsNumeric(mapa) Then
            X = ReadField(2, rdata, 32)
            Y = ReadField(3, rdata, 32)
            If MapaValido(mapa) = False Or InMapBounds(mapa, X, Y) = False Then
                Call SendData(ToIndex, UserIndex, 0, "||Mapa invalido." & FONTTYPE_INFO)
                Exit Function
            End If
        Else
            Select Case Left(UCase$(mapa), 3)
                Case "LIN"
                    mapa = Lindos.Map
                    X = Lindos.X
                    Y = Lindos.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Lindos." & FONTTYPE_INFO)
                Case "ULL"
                    mapa = Ullathorpe.Map
                    X = Ullathorpe.X
                    Y = Ullathorpe.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Ulla." & FONTTYPE_INFO)
                Case "BAN"
                    mapa = Banderbill.Map
                    X = Banderbill.X
                    Y = Banderbill.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Bander." & FONTTYPE_INFO)
                Case "NIX"
                    mapa = Nix.Map
                    X = Nix.X
                    Y = Nix.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Nix." & FONTTYPE_INFO)
                Case Else
                    Call SendData(ToIndex, UserIndex, 0, "||Mapa desconocido. Solo reconoce Lindos, Bander, Ulla y Nix." & FONTTYPE_INFO)
                Exit Function
            End Select
        End If
        MapData(MapaX, xx, Yx).TileExit.Map = mapa
        MapData(MapaX, xx, Yx).TileExit.X = X
        MapData(MapaX, xx, Yx).TileExit.Y = Y
        MapInfo(MapaX).Telep = True
        Call SendData(ToIndex, UserIndex, 0, "||Teletransporte Transparente creado." & FONTTYPE_INFO)
        Exit Function
    End If
    ' [/GS]
    
    
    'Crear Teleport
    If UCase(Left(rdata, 3)) = "/CT" Then
        '/ct mapa_dest x_dest y_dest
        If Len(rdata) < 4 Then
            Call SendData(ToIndex, UserIndex, 0, "||Comando mal utilizado. Syntaxis: /CT <Mapa> <X> <Y>" & FONTTYPE_INFO)
            Exit Function
        End If
        rdata = Right(rdata, Len(rdata) - 4)
        Call LogGM(UserList(UserIndex).Name, "/CT: " & rdata, False)
        mapa = ReadField(1, rdata, 32)
        If IsNumeric(mapa) Then
            X = ReadField(2, rdata, 32)
            Y = ReadField(3, rdata, 32)
            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Teleport no creado, no hay lugar." & FONTTYPE_INFO)
                Exit Function
            End If
            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Teleport no creado, lugar invalido." & FONTTYPE_INFO)
                Exit Function
            End If
            If MapaValido(mapa) = False Or InMapBounds(mapa, X, Y) = False Then
                Call SendData(ToIndex, UserIndex, 0, "||Mapa invalido." & FONTTYPE_INFO)
                Exit Function
            End If
        Else
            Select Case Left(UCase$(mapa), 3)
                Case "LIN"
                    mapa = Lindos.Map
                    X = Lindos.X
                    Y = Lindos.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Lindos." & FONTTYPE_INFO)
                Case "ULL"
                    mapa = Ullathorpe.Map
                    X = Ullathorpe.X
                    Y = Ullathorpe.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Ulla." & FONTTYPE_INFO)
                Case "BAN"
                    mapa = Banderbill.Map
                    X = Banderbill.X
                    Y = Banderbill.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Bander." & FONTTYPE_INFO)
                Case "NIX"
                    mapa = Nix.Map
                    X = Nix.X
                    Y = Nix.Y
                    Call SendData(ToIndex, UserIndex, 0, "||Teleport a Nix." & FONTTYPE_INFO)
                Case Else
                    Call SendData(ToIndex, UserIndex, 0, "||Mapa desconocido. Solo reconoce Lindos, Bander, Ulla y Nix." & FONTTYPE_INFO)
                Exit Function
            End Select
        End If
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1)
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map = mapa
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.X = X
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Y = Y
        Call SendData(ToIndex, UserIndex, 0, "||Teletransporte creado." & FONTTYPE_INFO)
        MapInfo(UserList(UserIndex).Pos.Map).Telep = True
        Exit Function
    End If
    
    'Destruir Teleport
    'toma el ultimo click
    If UCase(Left(rdata, 3)) = "/DT" Then
        '/dt
        Call LogGM(UserList(UserIndex).Name, "/DT", False)
        
        mapa = UserList(UserIndex).flags.TargetMap
        X = UserList(UserIndex).flags.TargetX
        Y = UserList(UserIndex).flags.TargetY
        
        If MapData(mapa, X, Y).TileExit.Map > 0 Then
            If ObjData(MapData(mapa, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT Then Call EraseObj(ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
            MapData(mapa, X, Y).TileExit.Map = 0
            MapData(mapa, X, Y).TileExit.X = 0
            MapData(mapa, X, Y).TileExit.Y = 0
            Call SendData(ToIndex, UserIndex, 0, "||Teletransporte eliminado." & FONTTYPE_INFO)
            MapInfo(mapa).Telep = True
        End If
        
        Exit Function
    End If

    
    'Destruir
    If UCase$(Left$(rdata, 5)) = "/DEST" Then
        Call C_Dest(UserIndex, rdata)
        Exit Function
    End If
    
    'Bloquear
    If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
        Call C_Bloq(UserIndex, rdata)
        Exit Function
    End If
    
    'Quitar NPC
    If UCase$(rdata) = "/MATA" Then
        Call C_Mata(UserIndex, rdata)
        Exit Function
    End If
    ' [GS]
    If UCase$(rdata) = "/BORRAR" Then
        rdata = Right$(rdata, Len(rdata) - 7)
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Function
        EraseObj ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY
        Call SendData(ToIndex, UserIndex, 0, "||Objeto borrado." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Mapa:" & UserList(UserIndex).Pos.Map & " /BORRAR " & ObjData(UserList(UserIndex).flags.TargetObj).Name, False)
        Exit Function
    End If
    ' [/GS]
    'Quita todos los NPCs del area
    If UCase$(rdata) = "/MASSKILL" Then
        Call C_MassKill(UserIndex, rdata)
        Exit Function
    End If
    
    'Quita todos los NPCs del area
    If UCase$(rdata) = "/MASSKILLGROX" Then
        For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                        If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex)
                    End If
                Next X
        Next Y
        SendData ToIndex, UserIndex, 0, "||Borrados todos los NPC del mapa entero." & FONTTYPE_FIGHT_YO
        Call LogGM(UserList(UserIndex).Name, "/MASSKILLGROX", False)
        Exit Function
    End If
    
    ' [NEW] Borra oro en el mapa
    If UCase$(rdata) = "/MASSDEST" Then
        Call C_MassDest(UserIndex, rdata)
        Exit Function
    End If
    ' [/NEW]

    
    'Mensaje del sistema
    If UCase$(Left$(rdata, 6)) = "/SMSG " Then
        Call C_Smsg(UserIndex, rdata)
        Exit Function
    End If
    
    'Crear criatura, toma directamente el indice
    If UCase$(Left$(rdata, 5)) = "/ACC " Then
        Call C_Acc(UserIndex, rdata)
        Exit Function
    End If
    
    'Crear criatura con respawn, toma directamente el indice
    If UCase$(Left$(rdata, 6)) = "/RACC " Then
        Call C_Racc(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AI1 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraImperial1 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AI2 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraImperial1 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AI3 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraImperial3 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AI4 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       TunicaMagoImperial = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AC1 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraCaos1 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AC2 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraCaos2 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AC3 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       ArmaduraCaos3 = val(rdata)
       Exit Function
    End If
    
    If UCase$(Left$(rdata, 5)) = "/AC4 " Then
       rdata = Right$(rdata, Len(rdata) - 5)
       TunicaMagoCaos = val(rdata)
       Exit Function
    End If
    
    
    
    'Comando para depurar la navegacion
    If UCase$(rdata) = "/NAVE" Then
        If UserList(UserIndex).flags.Navegando = 1 Then
            UserList(UserIndex).flags.Navegando = 0
        Else
            UserList(UserIndex).flags.Navegando = 1
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Depurando la navegacion." & FONTTYPE_INFO)
        Exit Function
    End If
    
    'Apagamos
    If UCase$(rdata) = "/APAGAR" Then
        If UCase$(UserList(UserIndex).Name) <> "GS" Then
            Call SendData(ToIndex, UserIndex, 0, "||Comando restringido." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
            Exit Function
        End If
        'Log
        mifile = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #mifile
        Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
        Close #mifile
        Unload frmGeneral
        Exit Function
    End If
    
    If Left$(UCase$(rdata), 5) = "/BAS " Then
        tIndex = NameIndex(Right$(rdata, Len(rdata) - 5))
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline o inexistente." & FONTTYPE_INFO)
            Exit Function
        End If
        If EsNewbie(tIndex) = True Then
            If UserList(tIndex).flags.BorrarAlSalir = False Then
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tIndex).Name & " marcado para ser borrado." & FONTTYPE_INFO)
                UserList(tIndex).flags.BorrarAlSalir = True
            Else
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tIndex).Name & " ya no sera borrado." & FONTTYPE_INFO)
                UserList(tIndex).flags.BorrarAlSalir = False
            End If
        Else
            If UserList(tIndex).flags.BorrarAlSalir = True Then
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tIndex).Name & " ya no sera borrado." & FONTTYPE_INFO)
                UserList(tIndex).flags.BorrarAlSalir = False
            Else
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tIndex).Name & " no es newbie, por lo tanto no se puede utilizar este comando." & FONTTYPE_INFO)
            End If
        End If
        Exit Function
    End If
    
    If Left$(UCase$(rdata), 10) = "/KILLCHAR " Then
        mifile = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #mifile
        Print #mifile, Date & " " & Time & Right$(rdata, Len(rdata) - 10) & " Borrado por " & UserList(UserIndex).Name & ". "
        Close #mifile
        rdata = Right$(rdata, Len(rdata) - 10)
        rdata = Replace(rdata, ".", " ")
        rdata = Replace(rdata, "+", " ")
        MatarPersonaje rdata
        Call SendData(ToIndex, UserIndex, 0, "||Personaje borrado." & FONTTYPE_INFO)
        Exit Function
    End If
    
    'CONDENA
    If UCase$(Left$(rdata, 7)) = "/CONDEN" Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(rdata)
        If tIndex > 0 Then
            Call VolverCriminal(tIndex)
            Call LogGM(UserList(UserIndex).Name, "Volvio criminal a " & rdata, False)
            Call SendData(ToIndex, UserIndex, 0, "||Personaje vuelto criminal." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        tIndex = NameIndex(UCase$(rdata))
        If tIndex > 0 Then
            Call ResetFacciones(tIndex)
            Call SendData(ToIndex, UserIndex, 0, "||Facciones reiniciadas." & FONTTYPE_INFO)
        End If
        Exit Function
    End If
    
    
    'MODIFICA CARACTER
    If UCase$(Left$(rdata, 5)) = "/MOD " Then
        Call LogGM(UserList(UserIndex).Name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 5)
        If ReadField(1, rdata, 32) = "yo" Then
            Name = UserList(UserIndex).Name
        Else
            Name = ReadField(1, rdata, 32)
        End If
        tIndex = NameIndex(Name)
        Arg1 = ReadField(2, rdata, 32)
        Arg2 = ReadField(3, rdata, 32)
        Arg3 = ReadField(4, rdata, 32)
        Arg4 = ReadField(5, rdata, 32)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Function
        ' [GS] Logs de editados
        ElseIf (UserList(tIndex).flags.Privilegios < 1 And EsAdmin(tIndex) = False) Then
            ' Si el editado no es GM
            Call LogCOSAS("Editados", UserList(UserIndex).Name & " edito a " & UserList(tIndex).Name & " : " & rdata, False)
        ' [/GS]
        End If
        
        Select Case UCase$(Arg1)
            ' v0.12a9
            Case "SKILL"
                If IsNumeric(Arg2) Then
                    If val(Arg2) <= 255 And val(Arg) >= 0 Then
                        Call ModSkills(tIndex, Arg2)
                        Call SendUserStatsBox(tIndex)
                        Call SendData(ToIndex, UserIndex, 0, "||Los Skilles de " & UserList(tIndex).Name & " han sido editados a " & Arg2 & "." & FONTTYPE_INFO)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor invalido." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                End If
                Exit Function
            ' v0.12a9
            Case "DADOS"
                If IsNumeric(Arg2) Then
                    If val(Arg2) <= 1024 And val(Arg2) > 0 Then
                        Call ModDados(tIndex, Arg2)
                        Call SendUserStatsBox(tIndex)
                        Call SendData(ToIndex, UserIndex, 0, "||Los Dados de " & UserList(tIndex).Name & " han sido editados a " & Arg2 & "." & FONTTYPE_INFO)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor invalido." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                End If
                Exit Function
                
            Case "ORO"
                'If val(Arg2) < 95001 Then
                If IsNumeric(Arg2) Then
                    If val(Arg2) > MaxOro Then Arg2 = MaxOro
                    UserList(tIndex).Stats.GLD = val(Arg2)
                    Call SendUserStatsBox(tIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Oro modificado." & FONTTYPE_INFO)
                ElseIf Left(UCase(Arg2), 3) = "MAX" Then
                    UserList(tIndex).Stats.GLD = MaxOro
                    Call SendUserStatsBox(tIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Oro al maximo." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                End If
                'Else
                '    Call SendData(ToIndex, UserIndex, 0, "||No esta permitido utilizar valores mayores a 95000. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                Exit Function
                'End If
            Case "EXP"
                'If val(Arg2) < 9995001 Then
                 If IsNumeric(Arg2) Then
                    If UserList(tIndex).Stats.exp + val(Arg2) > UserList(tIndex).Stats.ELU Then
                       'Dim resto
                       'resto = val(Arg2) - UserList(tIndex).Stats.ELU
                       UserList(tIndex).Stats.exp = UserList(tIndex).Stats.exp + UserList(tIndex).Stats.ELU
                       'Call CheckUserLevel(tIndex)
                       'UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                       'Call CheckUserLevel(tIndex)
                       Call SendData(ToIndex, UserIndex, 0, "||Experiencia modificada." & FONTTYPE_INFO)
                    Else
                       UserList(tIndex).Stats.exp = val(Arg2)
                    End If
                ElseIf Left(UCase(Arg2), 3) = "MAX" Then
                    UserList(tIndex).Stats.exp = UserList(tIndex).Stats.ELU + 1
                    Call SendData(ToIndex, UserIndex, 0, "||Exp al maximo." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                    Exit Function
                End If
                Call CheckUserLevel(tIndex)
                Call SendUserStatsBox(tIndex)
                Exit Function
            Case "HIT"
                If IsNumeric(Arg2) Then
                    If Arg2 > STAT_MAXHIT Then
                        UserList(tIndex).Stats.MaxHIT = STAT_MAXHIT
                        Call SendData(ToIndex, UserIndex, 0, "||HIT al maximo." & FONTTYPE_INFO)
                    Else
                        UserList(tIndex).Stats.MaxHIT = val(Arg2)
                        Call SendData(ToIndex, UserIndex, 0, "||HIT modificado." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                End If
                Exit Function
            Case "DEF"
                If IsNumeric(Arg2) Then
                    If Arg2 > STAT_MAXDEF Then
                        UserList(tIndex).Stats.Def = STAT_MAXDEF
                        Call SendData(ToIndex, UserIndex, 0, "||DEF al maximo." & FONTTYPE_INFO)
                    Else
                        UserList(tIndex).Stats.Def = val(Arg2)
                        Call SendData(ToIndex, UserIndex, 0, "||DEF modificado." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: Valor no reconocido." & FONTTYPE_INFO)
                End If
                Exit Function
            Case "BODY"
                Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Call SendData(ToIndex, UserIndex, 0, "||Cuerpo cambiado." & FONTTYPE_INFO)
                Exit Function
            Case "HEAD"
                Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                If UserIndex = tIndex And val(Arg2) > 0 Then
                    UserList(tIndex).Char.Head = val(Arg2)
                    UserList(tIndex).OrigChar.Head = val(Arg2)
                    Call SendData(ToIndex, UserIndex, 0, "||Cabeza cambiada y guardada." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Cabeza cambiada." & FONTTYPE_INFO)
                End If
                Exit Function
            Case "NICK"
                If UCase(Arg2) = "GS" Then
                    Call SendData(ToIndex, UserIndex, 0, "||Nick reservado." & FONTTYPE_INFO)
                    Exit Function
                End If
                If Len(Arg2) < 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||El nombre debe tener menos de 2 letras." & FONTTYPE_INFO)
                    Exit Function
                End If
                If Len(Arg2) > 30 Then
                    Call SendData(ToIndex, UserIndex, 0, "||El nombre debe tener como maximo de 30 letras." & FONTTYPE_INFO)
                    Exit Function
                End If
                'If Not NombrePermitido(str(Arg2)) And Inbaneable(str(Arg2)) = False Then
                '    Call SendData(ToIndex, Userindex, 0, "||El nombre indicado es invalido. Posiblemente no sea apropiado para todo el publico." & FONTTYPE_INFO)
                '    Exit Function
                'End If
                
                Arg2 = Replace(Arg2, "+", " ")
                Arg2 = Replace(Arg2, ".", " ")
                
                If Not AsciiValidos(Arg2) Then
                    Call SendData(ToIndex, UserIndex, 0, "||Nombre invalido. Tiene caracteres invalidos." & FONTTYPE_INFO)
                    Exit Function
                End If
                If FileExist(CharPath & UCase$(Arg2) & ".chr", vbNormal) = True Then
                    Call SendData(ToIndex, UserIndex, 0, "||Ya existe el personaje '" & UCase(Arg2) & "'." & FONTTYPE_INFO)
                    Exit Function
                End If
                MatarPersonaje UserList(tIndex).Name
                UserList(tIndex).Name = UCase(Arg2)
                Call SendData(ToIndex, UserIndex, 0, "||El Nombre del Personaje ha sido cambiado." & FONTTYPE_INFO)
                Call SendData(ToIndex, tIndex, 0, "||Tú Nombre ha sido cambiado a '" & UCase(Arg2) & "'" & FONTTYPE_INFO)
                ' [GS] Corrige error de mapa
                Call ResetUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex)
                ' [/GS]
            Case "EMAIL"
                UserList(tIndex).Email = Arg2
                Call SendData(ToIndex, UserIndex, 0, "||E-mail cambiado." & FONTTYPE_INFO)
                Call SendData(ToIndex, tIndex, 0, "||Tú E-mail ha sido cambiado a " & Arg2 & "" & FONTTYPE_INFO)
                Exit Function
            Case "CRI"
                UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
                Call SendData(ToIndex, UserIndex, 0, "||Criminales Matados cambiados." & FONTTYPE_INFO)
                Exit Function
            Case "CIU"
                UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
                Call SendData(ToIndex, UserIndex, 0, "||Ciudadanos Matados cambiados." & FONTTYPE_INFO)
                Exit Function
            Case "LEVEL"
            If val(Arg2) <= STAT_MAXELV Then
                UserList(tIndex).Stats.ELV = val(Arg2)
                Call SendData(ToIndex, UserIndex, 0, "||Nivel cambiado." & FONTTYPE_INFO)
                Exit Function
            Else
                Call SendData(ToIndex, UserIndex, 0, "||MOD: No esta permitido utilizar valores mayores a " & STAT_MAXELV & "." & FONTTYPE_INFO)
                Exit Function
            End If
            Case "GEN"
                If Not Numeric(val(Arg2)) Then Exit Function
                UserList(tIndex).genero = IIf(val(Arg2) = 2, MUJER, HOMBRE)
                Call SendUserStatsBox(tIndex)
                Call SendData(ToIndex, UserIndex, 0, "||Genero cambiado." & FONTTYPE_INFO)
                Exit Function
                ' [NEW]
             Case "HP"
                If val(Arg2) <= STAT_MAXHP Then
                    UserList(tIndex).Stats.MinHP = val(Arg2)
                    UserList(tIndex).Stats.MaxHP = val(Arg2)
                    Call SendUserStatsBox(tIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Vida cambiada." & FONTTYPE_INFO)
                    'Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " le edito la vida hacia " & val(Arg2) & " a " & UserList(tIndex).Name & ".- Escrachado, eh? :P" & FONTTYPE_FIGHT)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: No esta permitido utilizar valores mayores a " & STAT_MAXHP & ". Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                    Exit Function
                End If
                Exit Function
            Case "ST"
                If val(Arg2) < 50000 Then
                    UserList(tIndex).Stats.MinSta = val(Arg2)
                    UserList(tIndex).Stats.MaxSta = val(Arg2)
                    Call SendUserStatsBox(tIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Stamina cambiada." & FONTTYPE_INFO)
                    'Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " le edito el mana hacia " & val(Arg2) & " a " & UserList(tIndex).Name & ".- Escrachado, eh? :P" & FONTTYPE_FIGHT)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: No esta permitido utilizar valores mayores a 50000. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                    Exit Function
                End If
                Exit Function
                ' [/NEW]
            Case "MP"
                If val(Arg2) < 50000 Then
                    UserList(tIndex).Stats.MinMAN = val(Arg2)
                    UserList(tIndex).Stats.MaxMAN = val(Arg2)
                    Call SendUserStatsBox(tIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Mana cambiada." & FONTTYPE_INFO)
                    'Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " le edito el mana hacia " & val(Arg2) & " a " & UserList(tIndex).Name & ".- Escrachado, eh? :P" & FONTTYPE_FIGHT)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||MOD: No esta permitido utilizar valores mayores a 50000. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                    Exit Function
                End If
                Exit Function
                ' [/NEW]
            Case Else
                Call SendData(ToIndex, UserIndex, 0, "||MOD: Comando no existente." & FONTTYPE_INFO)
                Exit Function
        End Select
    
        Exit Function
    End If
    
    
    If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
        If haciendoBK = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya se esta realizando un backup" & FONTTYPE_INFO)
            Exit Function
        End If
        Call DoBackUp
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 7)) = "/GRABAR" Then
        If haciendoBK = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya se esta realizando un backup." & FONTTYPE_INFO)
            Exit Function
        End If
        Call GuardarUsuarios
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 7)) = "/SEGURO" Then
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
            MapInfo(UserList(UserIndex).Pos.Map).Pk = False
            Call SendData(ToIndex, UserIndex, 0, "||Ahora es zona segura." & FONTTYPE_INFO)
            Exit Function
        Else
            MapInfo(UserList(UserIndex).Pos.Map).Pk = True
            Call SendData(ToIndex, UserIndex, 0, "||Ahora es zona insegura." & FONTTYPE_INFO)
            Exit Function
        End If
        MapInfo(UserList(UserIndex).Pos.Map).Datos = True
        Exit Function
    End If
    
    ' [GS]
    If UCase$(rdata) = "/BACK" Then
        If haciendoBK = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya se esta realizando un backup" & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(UserIndex).Pos.Map = MAPA_PRETORIANO Then
            Call SendData(ToIndex, UserIndex, 0, "||No esta permitido hacer un /BACK de este mapa, con el sistema Pretoriano Activado." & FONTTYPE_INFO)
            Exit Function
        End If
        haciendoBK = True
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToAll, 0, 0, "||Guardando mapa " & UserList(UserIndex).Pos.Map & "..." & FONTTYPE_INFO)
        SaveMapData UserList(UserIndex).Pos.Map
        Call SendData(ToAll, 0, 0, "||Mapa guardado." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        haciendoBK = False
        Exit Function
    End If
    ' [/GS]
    
    ' [GS]
    If UCase$(rdata) = "/BACKNPC" Then
        If haciendoBK = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya se esta realizando un backup" & FONTTYPE_INFO)
            Exit Function
        End If
        haciendoBK = True
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToAll, 0, 0, "||Guardando NPC's..." & FONTTYPE_INFO)
        If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
        If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")
        For LoopC = 1 To LastNPC
            If Npclist(LoopC).flags.BackUp = 1 Then
                    Call BackUPnPc(LoopC)
            End If
        Next
        Call SendData(ToAll, 0, 0, "||NPC's guardados." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        haciendoBK = False
        Exit Function
    End If
    ' [/GS]
    
    ' [GS]
    If UCase$(rdata) = "/NIVELES" Then
        tStr = ""
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).NoExiste = False Then
                tStr = tStr & UserList(LoopC).Name & "(" & UserList(LoopC).Stats.ELV & "), "
            End If
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Exit Function
    End If
    
    ' [/GS]
    
    ' [GS] AYUDANTES
    If UCase$(Left$(rdata, 10)) = "/AYUDANTE " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        If rdata <> "" Then
            Name = rdata
            If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
                If EsAyudante(Name) = False Then
                    If NameIndex(Name) > 0 Then
                        Call SendData(ToAyudantes, 0, 0, "!!Bienvenida '" & Name & "', ahora eres Ayudante." & FONTTYPE_ADMIN)
                        Call SendData(ToIndex, UserIndex, 0, "FINOK")
                        Call CloseUser(UserIndex)
                    End If
                    Call SendData(ToAyudantes, 0, 0, "||Le damos la Bienvenida al nuevo ayudante '" & Name & "'." & FONTTYPE_ADMIN)
                    Call PonerAyudante(Name)
                Else
                    If NameIndex(Name) > 0 Then
                        Call SendData(ToAyudantes, 0, 0, "!!Has sido expulsado de los Ayudantes." & FONTTYPE_ADMIN)
                        Call SendData(ToIndex, UserIndex, 0, "FINOK")
                        Call CloseUser(UserIndex)
                    End If
                    Call SendData(ToAyudantes, 0, 0, "||El Ayudante '" & Name & "' ha sido relevado de su cargo." & FONTTYPE_ADMIN)
                    Call QuitarAyudante(Name)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No existe el personaje." & FONTTYPE_INFX)
            End If
        End If
        Exit Function
    End If
    ' [/GS]
    
    ' [GS]
    If UCase$(rdata) = "/RES" Then
        Call SendData(ToIndex, UserIndex, 0, "||Modo de uso:" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||/RES <tipo> <pots> <no-ko> <NoSeCaenItems>" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||1ro. Tipo:" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||    1 - Vate TODO" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||    2 - No vale Ceguera, Estupides y Invi" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||    3 - Idem pero sin Para o Inmo" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||    4 - Torneo sin magias, solo a Cuchi" & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||2do. Pots: 1-Si se permiten  0-Si no " & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||3ro. Maximo de Mascotas el torneo." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||4to. Si existen ataques que maten de una, 1-No, 0-Si." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||5to. Inidica si se caen los objetos al morir, 1-No, 0-Si." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Ejemplo: /RES 2 1 0 1 1" & FONTTYPE_INFO)
        Exit Function
    End If
    If UCase$(Left$(rdata, 5)) = "/RES " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Arg1 = ReadField(1, rdata, 32) ' tipo
        Arg2 = ReadField(2, rdata, 32) ' pots
        Arg3 = ReadField(3, rdata, 32) ' mascotas
        Arg4 = ReadField(4, rdata, 32) ' no ko?
        Arg5 = ReadField(5, rdata, 32) ' NoSeCaen
            Select Case Arg1
                Case 1
                    Call SendData(ToAll, 0, 0, "||<Torneo> Vate TODO" & FONTTYPE_TALK & ENDC)
                Case 2
                    Call SendData(ToAll, 0, 0, "||<Torneo> No vale Ceguera, Estupides y Invi" & FONTTYPE_TALK & ENDC)
                Case 3
                    Call SendData(ToAll, 0, 0, "||<Torneo> No vale Ceguera, Estupides, Invi, Para o Inmo" & FONTTYPE_TALK & ENDC)
                Case 4
                    Call SendData(ToAll, 0, 0, "||<Torneo> No vale ningun hechizo, solo es a Cuchi" & FONTTYPE_TALK & ENDC)
                Case Else
                    Call SendData(ToIndex, UserIndex, 0, "||Tipo incorrecto, solo 1, 2, 3 o 4... Coloca /RES, sin parametros para ver como utilizarlo." & FONTTYPE_INFO)
                    Exit Function
            End Select
            ConfigTorneo = Arg1
            If Arg2 <> 0 Then
                PotsEnTorneo = True
                Call SendData(ToAll, 0, 0, "||<Torneo> Las pots estan permitidas en el torneo." & FONTTYPE_TALK & ENDC)
            Else
                PotsEnTorneo = False
                Call SendData(ToAll, 0, 0, "||<Torneo> Las pots no estan permitidas en el torneo." & FONTTYPE_TALK & ENDC)
            End If
            If Arg3 > MAXMASCOTAS Then
                MaxMascotasTorneo = MAXMASCOTAS
                Call SendData(ToAll, 0, 0, "||<Torneo> Solo se permiten " & Arg3 & " mascotas." & FONTTYPE_TALK & ENDC)
            ElseIf Arg3 < 1 Then
                MaxMascotasTorneo = 0
                Call SendData(ToAll, 0, 0, "||<Torneo> Las mascotas no estan permitidas." & FONTTYPE_TALK & ENDC)
            ElseIf IsNumeric(Arg3) Then
                MaxMascotasTorneo = Arg3
                Call SendData(ToAll, 0, 0, "||<Torneo> Solo se permiten " & Arg3 & " mascotas." & FONTTYPE_TALK & ENDC)
            Else
                MaxMascotasTorneo = 0
                Call SendData(ToAll, 0, 0, "||<Torneo> Las mascotas no estan permitidas." & FONTTYPE_TALK & ENDC)
            End If
            If Arg4 <> 0 Then
                NoKO = True
                Call SendData(ToAll, 0, 0, "||<Torneo> Nadie podra matar de un golpe a su oponente." & FONTTYPE_TALK & ENDC)
            Else
                NoKO = False
                Call SendData(ToAll, 0, 0, "||<Torneo> Se podra matar de un golpe al oponente." & FONTTYPE_TALK & ENDC)
            End If
            If Arg5 <> 0 Then
                NoSeCaenItemsEnTorneo = True
                Call SendData(ToAll, 0, 0, "||<Torneo> No se caeran los items al morir." & FONTTYPE_TALK & ENDC)
            Else
                NoSeCaenItemsEnTorneo = False
                Call SendData(ToAll, 0, 0, "||<Torneo> Los items se caeran al morir." & FONTTYPE_TALK & ENDC)
            End If
            ' Guarda en el INI la configuracion
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "ConfigTorneo", val(Arg1))
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "ValenPots", IIf(PotsEnTorneo = True, "1", "0"))
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "MaxMascotasTorneo", val(Arg3))
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "NoKO", IIf(NoKO = True, "1", "0"))
            Call WriteVar(App.Path & "\Opciones.ini", "TORNEO", "NoSeCaenItemsEnTorneo", IIf(NoSeCaenItemsEnTorneo = True, "1", "0"))
            Call SendData(ToIndex, UserIndex, 0, "||Configuración guardada..." & FONTTYPE_INFO)
        Exit Function
    End If
    '[/GS]
    ' ### SETEA EL MAPA DE TORNEO ###
    
    If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
        Call Ayuda.Reset
        Call SendData(ToIndex, UserIndex, 0, "||SOS reseteado." & FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
        ' [GS]
        'Call frmGeneral.mnuMostrar_Click
        'Call frmGeneral.mnuMostrar
        frmGeneral.QuitarSysTray
        frmGeneral.WindowState = vbNormal
        frmGeneral.Visible = True
    '    frmGeneral.WindowState = 0  ' Ventana normal
    '    frmGeneral.Visible = True   ' y visible
        Call SendData(ToIndex, UserIndex, 0, "||General visible..." & FONTTYPE_INFO)
        ' [/GS]
        Exit Function
    End If
    
    If UCase$(rdata) = "/LLUVIA" Then
        Call C_Lluvia(rdata)
        Exit Function
    End If
    
    If UCase$(rdata) = "/PASSDAY" Then
        Call DayElapsed
        Exit Function
    End If
    
    If UCase$(rdata) = "/PRETORIAN" Then
        If MAPA_PRETORIANO <> 0 Then ' Solo si usamos modo pretorian aplicamos
        ' [EL OSO]
            Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
        ' [/EL OSO]
            Call SendData(ToIndex, UserIndex, 0, "||Clan Pretoriano Creado..." & FONTTYPE_INFX)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No se encuentra especificado el Mapa Pretoriano..." & FONTTYPE_INFX)
        End If
    End If
    
    If UCase$(rdata) = "/HABILITAR" Then
        If ReservadoParaAdministradores = True Then
            ReservadoParaAdministradores = False
            Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "ReservadoParaAdministradores", 0)
            Call SendData(ToIndex, UserIndex, 0, "||El servidor esta habilitado para todo publico." & FONTTYPE_ADMIN)
        Else
            ReservadoParaAdministradores = True
            Call WriteVar(IniPath & "Opciones.ini", "SERVIDOR", "ReservadoParaAdministradores", 1)
            Call SendData(ToIndex, UserIndex, 0, "||El servidor esta reservado para administradores." & FONTTYPE_ADMIN)
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") Then
                    If UserList(LoopC).flags.Privilegios > 0 Or EsAdmin(LoopC) Then ' Es GM
                    
                    Else
                        Call SendData(ToIndex, LoopC, 0, "ERREl servidor esta reservado solo para Administradores" & IIf(Len(URL_Soporte) > 2, ". Mas información " & URL_Soporte, "."))
                        Call SendData(ToIndex, LoopC, 0, "FINOK")
                        Call CloseUser(LoopC)
                    End If
                End If
            Next LoopC
            Call SendData(ToIndex, UserIndex, 0, "||Se han desconectado todos los usuarios sin privilegios." & FONTTYPE_ADMIN)
        End If
        Exit Function
    End If
    
    If UCase$(rdata) = "/MASMANTENIMIENTO" Then
        If HsMantenimiento < 60 Then
            HsMantenimiento = HsMantenimiento + 60
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ha agregado 1 Hora al tiempo restante del Mantenimiento." & FONTTYPE_ADMIN)
        End If
        Exit Function
    End If
    
    ' 0.12b1
    '[yb]
    If UCase$(Left$(rdata, 12)) = "/ACEPTCONSE " Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Function
        rdata = Right$(rdata, Len(rdata) - 12)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el honorable Consejo Real de Banderbill." & FONTTYPE_CONSEJO)
            UserList(tIndex).flags.PertAlCons = 1
            Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
        End If
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 16)) = "/ACEPTCONSECAOS " Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Function
        rdata = Right$(rdata, Len(rdata) - 16)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el Consejo de la Legión Oscura." & FONTTYPE_CONSEJOCAOS)
            UserList(tIndex).flags.PertAlConsCaos = 1
            Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
        End If
        Exit Function
    End If
    
    
    
    If Left$(UCase$(rdata), 5) = "/PISO" Then
        Call C_Piso(UserIndex, rdata)
        Exit Function
    End If
    
    If UCase$(Left$(rdata, 11)) = "/KICKCONSE " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        tIndex = NameIndex(rdata)
        If tIndex <= 0 Then
            If FileExist(CharPath & rdata & ".chr", vbArchive) Then
                Call SendData(ToIndex, UserIndex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
                Call WriteVar(CharPath & UCase(rdata) & ".chr", "CONSEJO", "PERTENECE", 0)
                Call WriteVar(CharPath & UCase(rdata) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No se encuentra el charfile " & CharPath & rdata & ".chr" & FONTTYPE_INFO)
                Exit Function
            End If
        Else
            If UserList(tIndex).flags.PertAlCons > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
                UserList(tIndex).flags.PertAlCons = 0
                Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)
            End If
            If UserList(tIndex).flags.PertAlConsCaos > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
                UserList(tIndex).flags.PertAlConsCaos = 0
                Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)
            End If
        End If
        Exit Function
    End If
    '[/yb]
    
    ' 0.12b1
    If UCase$(Left$(rdata, 8)) = "/NOCAOS " Then
        Call C_NoCaos(UserIndex, rdata)
        Exit Function
    End If
    If UCase$(Left$(rdata, 8)) = "/NOREAL " Then
        Call C_NoReal(UserIndex, rdata)
        Exit Function
    End If
    
    TCP_Admin = False
End Function
