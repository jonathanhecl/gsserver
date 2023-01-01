Attribute VB_Name = "GS_NPC_Agite"
Sub aNPC_Pensar(ByVal NpcIndex As Integer)
On Error GoTo fallo
    Dim Amigo As Integer
    Dim HayCosas As String
    Dim RecienHablo As Boolean
    'Dim NUMX As Integer
    'NUMX = CInt(RandomNumber(1, 100)) + 1

    Amigo = aNPC_Amigo(NpcIndex)
    HayCosas = aNPC_Cosas(NpcIndex)
    RecienHablo = False
    ' Que nuestra lokura nos acompañe siempre que veamos un usuario
    If aNPC_ALERTA(NpcIndex) = True Then
       ' If Amigo > 0 Then
       '     If Npclist(NpcIndex).flags.Dijo <> 5 Then
       '         Npclist(NpcIndex).flags.Dijo = 5
       '         Call aNPC_Habla(NpcIndex, "ayuda amigo")
       '     End If
       ' Else
       '     If Npclist(NpcIndex).flags.Dijo <> 4 Then
       '         Npclist(NpcIndex).flags.Dijo = 4
       '         Call aNPC_Habla(NpcIndex, "maldito nw")
       '     End If
       ' End If
        Call aNPC_Mover(NpcIndex, True)
    Else
        If aNPC_TienePOTSqNecesita(NpcIndex) Then
        '    If Npclist(NpcIndex).flags.Dijo <> 3 Then
        '        Npclist(NpcIndex).flags.Dijo = 3
        '        Call aNPC_Habla(NpcIndex, "j0z no paso el dopping")
        '        RecienHablo = True
        '    End If
            Call aNPC_UsamosPots(NpcIndex)
        ElseIf aNPC_TieneAlgunDaño(NpcIndex) Then
            Call aNPC_Curarse(NpcIndex)
        End If
    End If
    
    If Amigo > 0 Then
        'If Npclist(NpcIndex).flags.Dijo <> 1 Then
        '    Npclist(NpcIndex).flags.Dijo = 1
        '    Call aNPC_Habla(NpcIndex, "ya voy a ayudarte")
        'End If
        
        ' Tiene un amigo en problemas
        'Call ReCalculatePath(NpcIndex)
        Npclist(NpcIndex).PFINFO.Target.X = Npclist(Amigo).Pos.Y
        Npclist(NpcIndex).PFINFO.Target.Y = Npclist(Amigo).Pos.X
        Call aNPC_Mover(NpcIndex, False)
        If Npclist(Amigo).Stats.MinHP < (Npclist(Amigo).Stats.MaxHP / 2) Then
            Call aNPC_CurarAmigo(NpcIndex, Amigo)
        End If
    ElseIf Len(HayCosas) > 0 Then
        'If Npclist(NpcIndex).flags.Dijo <> 2 Then
        '    Npclist(NpcIndex).flags.Dijo = 2
        '    Call aNPC_Habla(NpcIndex, "pachanga dijo la changa")
        'End If
        'Call ReCalculatePath(NpcIndex)
        Dim X As Integer
        Dim Y As Integer
        X = CInt(ReadField(1, HayCosas, Asc("-"))) '- 1
        Y = CInt(ReadField(2, HayCosas, Asc("-"))) '+ 1
        Npclist(NpcIndex).PFINFO.Target.X = X
        Npclist(NpcIndex).PFINFO.Target.Y = Y
        Npclist(NpcIndex).PFINFO.TargetUser = 0
        'Call SendData(ToAll, 0, 0, "||" & Npclist(NpcIndex).PFINFO.Target.X & "-" & Npclist(NpcIndex).PFINFO.Target.Y & FONTTYPE_GS)
        Call aNPC_Mover(NpcIndex, False)
        If Npclist(NpcIndex).PFINFO.Target.X = Npclist(NpcIndex).Pos.X And Npclist(NpcIndex).PFINFO.Target.Y = Npclist(NpcIndex).Pos.Y Then Call aNPC_Agarrar(NpcIndex)                     ' si tengo algo debajo lo agarramos
    End If
    If RecienHablo = True Then Exit Sub
    If Npclist(NpcIndex).Meditando = True Then Exit Sub
    'If Npclist(NpcIndex).flags.Dijo <> 6 Then
    '    Npclist(NpcIndex).flags.Dijo = 6
    '    Call aNPC_Habla(NpcIndex, "la la la xD")
    'End If
Exit Sub

fallo:

'Call LogCOSAS("NPC de Agite", "ERROR " & Err.Number & " - " & Err.Description)

End Sub

Sub aNPC_Habla(ByVal NpcIndex As Integer, ByVal Dialogo As String)
Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "||" & vbWhite & "°" & Dialogo & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
Npclist(NpcIndex).flags.Hablo = True
End Sub

Sub aNPC_Curarse(ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
        'If Npclist(NpcIndex).flags.Dijo <> 8 Then
        '    Npclist(NpcIndex).flags.Dijo = 8
        '    Call aNPC_Habla(NpcIndex, "me lastime mucho")
        'End If
        Call aNPC_CurarAmigo(NpcIndex)
    ElseIf Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
        'If Npclist(NpcIndex).flags.Dijo <> 7 Then
        '    Npclist(NpcIndex).flags.Dijo = 7
        '    Call aNPC_Habla(NpcIndex, "esa pelea me canso")
        'End If
        Call NPCMeditando(NpcIndex, True)
    End If
End Sub

Function aNPC_TieneAlgunDaño(ByVal NpcIndex As Integer) As Boolean
    aNPC_TieneAlgunDaño = False
    If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
        aNPC_TieneAlgunDaño = True
    End If
    If Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
        aNPC_TieneAlgunDaño = True
    End If
End Function

Function aNPC_TienePOTSqNecesita(ByVal NpcIndex As Integer) As Boolean
On Error GoTo Error
    aNPC_TienePOTSqNecesita = False
    Dim NRojas As Boolean
    Dim NAzules As Boolean
    NRojas = False
    NAzules = False
    If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
        NRojas = True
    End If
    If Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
        NAzules = True
    End If
    If NRojas = False And NAzules = False Then Exit Function
    Dim Slot As Integer
    For Slot = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then
            If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).TipoPocion = 3 And NRojas = True Then aNPC_TienePOTSqNecesita = True
            If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).TipoPocion = 4 And NAzules = True Then aNPC_TienePOTSqNecesita = True
            If aNPC_TienePOTSqNecesita = True Then Exit Function
        End If
    Next
    aNPC_TienePOTSqNecesita = False
Exit Function
Error:
End Function

Sub aNPC_UsamosPots(ByVal NpcIndex As Integer)
    Dim Slot As Integer
    For Slot = 1 To MAX_INVENTORY_SLOTS
            ' Si ya tengo el objeto, vemos si nos sirve de algo
            'If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).MinHP > 0 Then
            If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).TipoPocion = 3 Then
                ' Nos cura
                If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
                    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + CInt(RandomNumber(ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).MinModificador, ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).MaxModificador))
                    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
                    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & SND_BEBER)
                    Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - 1
                    If Npclist(NpcIndex).Invent.Object(Slot).Amount < 1 Then
                        Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
                        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
                    End If
                    Exit Sub
                End If
            End If
            If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).TipoPocion = 4 Then
                ' Nos llena
                If Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
                    Npclist(NpcIndex).MiMana = Npclist(NpcIndex).MiMana + CInt(RandomNumber(ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).MinModificador, ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).MaxModificador))
                    If Npclist(NpcIndex).MiMana > Npclist(NpcIndex).mana Then Npclist(NpcIndex).MiMana = Npclist(NpcIndex).mana
                    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & SND_BEBER)
                    Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - 1
                    If Npclist(NpcIndex).Invent.Object(Slot).Amount < 1 Then
                        Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
                        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
                    End If
                End If
                Exit Sub
            End If
    Next

End Sub

Sub aNPC_CurarAmigo(ByVal NpcIndex As Integer, Optional Amigo As Integer)
    
    If Amigo = 0 Then Amigo = NpcIndex
    If Npclist(Amigo).Stats.MinHP = Npclist(Amigo).Stats.MaxHP Then Exit Sub ' No te curo viteh, porque no tiene gracia
    
    Dim i As Integer
    Dim MaxS As Integer
    Dim MaxI As Integer
    MaxS = 0
    MaxI = 0
    For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
        If Hechizos(Npclist(NpcIndex).Spells(i)).MinHP > 0 And Hechizos(Npclist(NpcIndex).Spells(i)).SubeHP = 1 Then
            If Hechizos(Npclist(NpcIndex).Spells(i)).MinHP > MaxS Then
                MaxS = Hechizos(Npclist(NpcIndex).Spells(i)).MinHP
                MaxI = Npclist(NpcIndex).Spells(i)
            End If
        End If
    Next
    If MaxI = 0 Then Exit Sub ' No tiene cura viteh
    'If Npclist(NpcIndex).MiMana < 50 Then
    '    Call aNPC_Habla(Amigo, "hay voy por tu ayuda")
    '    Call NPCMeditando(NpcIndex, True)
    '    Exit Sub
    'End If
    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "||" & vbCyan & "° " & Hechizos(MaxI).PalabrasMagicas & " °" & str(Npclist(NpcIndex).Char.CharIndex))
    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & Hechizos(MaxI).WAV)
    Call SendData(ToNPCArea, Amigo, Npclist(Amigo).Pos.Map, "CFX" & Npclist(Amigo).Char.CharIndex & "," & Hechizos(MaxI).FXgrh & "," & Hechizos(MaxI).loops)
    Npclist(NpcIndex).MiMana = Npclist(NpcIndex).MiMana - Hechizos(MaxS).ManaRequerido
    If Npclist(Amigo).Stats.MinHP + Hechizos(MaxS).MaxHP > Npclist(Amigo).Stats.MaxHP Then
        Npclist(Amigo).Stats.MinHP = Npclist(Amigo).Stats.MaxHP
    Else
        Npclist(Amigo).Stats.MinHP = Npclist(Amigo).Stats.MinHP + Hechizos(MaxS).MaxHP
    End If
    'If Amigo <> NpcIndex Then
    '    Call aNPC_Habla(Amigo, "gracias compañero")
    '    Call NPCMeditando(Amigo, True)
    'End If
    'Call NPCMeditando(NpcIndex, True)
End Sub

Sub aNPC_Agarrar(ByVal NpcIndex As Integer)
On Error Resume Next
    Dim Map, X, Y As Integer
    Dim Slot As Integer
    Map = Npclist(NpcIndex).Pos.Map
    X = Npclist(NpcIndex).Pos.X
    Y = Npclist(NpcIndex).Pos.Y
    If MapData(Map, X, Y).OBJInfo.ObjIndex < 0 Then Exit Sub
    'Sino se fija por un slot vacio antes del slot devuelto
    For Slot = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = (MapData(Map, X, Y).OBJInfo.ObjIndex) Then Exit For
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit For
    Next
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If Slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
        Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).TipoPocion = 3 Or ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).TipoPocion = 4 Then
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = MapData(Map, X, Y).OBJInfo.ObjIndex
            Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount + MapData(Map, X, Y).OBJInfo.Amount
            If Npclist(NpcIndex).Invent.Object(Slot).Amount > 10000 Then Npclist(NpcIndex).Invent.Object(Slot).Amount = 10000
            Call EraseObj(ToMap, 0, Npclist(NpcIndex).Pos.Map, MapData(Map, X, Y).OBJInfo.Amount, Map, X, Y)
        End If
        
    End If

End Sub

Sub aNPC_Mover(ByVal NpcIndex As Integer, Optional Agitando As Boolean)
On Error Resume Next
    ' Nos movemos alocadamente
    If Agitando = True Then
        Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
        If ReCalculatePath(NpcIndex) Then
            Call aNPC_PathFinding(NpcIndex)
            'Existe el camino?
            If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                'Move randomly
                Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
            End If
        Else
            If Not PathEnd(NpcIndex) Then
                Call FollowPath(NpcIndex)
            Else
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        End If
    Else
        Dim EstoyX, EstoyY As Integer
        EstoyX = Npclist(NpcIndex).Pos.X
        EstoyY = Npclist(NpcIndex).Pos.Y
        'Call SendData(ToAll, 0, 0, "||Voy a " & Npclist(NpcIndex).PFINFO.Target.X & "-" & Npclist(NpcIndex).PFINFO.Target.Y & FONTTYPE_GS)
        Call SeekPath(NpcIndex)
        Call FollowPath(NpcIndex)
        'Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
        If EstoyX = Npclist(NpcIndex).Pos.X And EstoyY = Npclist(NpcIndex).Pos.Y Then
            Dim TheyPos As WorldPos
            TheyPos.X = Npclist(NpcIndex).PFINFO.Target.X
            TheyPos.Y = Npclist(NpcIndex).PFINFO.Target.Y
            TheyPos.Map = Npclist(NpcIndex).Pos.Map
            tHeading = FindDirection(Npclist(NpcIndex).Pos, TheyPos)
            MoveNPCChar NpcIndex, tHeading
        End If
    End If
End Sub

Function aNPC_Cosas(ByVal NpcIndex As Integer) As String
    Dim Y, X, Map As Integer
    Dim NVida, NMana As Integer
    Map = Npclist(NpcIndex).Pos.Map
    aNPC_Cosas = ""
    NVida = 0
    NMana = 0
    ' Necesitades
    If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
        ' Toy sano, no necesito tanto pociones rojas
        NVida = 1
    End If
    If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
        ' Estoy herido, necesito pociones rojas
        NVida = 2
    End If
    If Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP / 2) Then
        ' Estoy MUY herido, necesito pociones rojas URGENTE
        NVida = 3
    End If
    If Npclist(NpcIndex).MiMana = Npclist(NpcIndex).mana Then
        ' Toy lleno, no necesito tanto pociones azules
        NMana = 1
    End If
    If Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana Then
        ' Estoy cansado, necesito pociones azules
        NMana = 2
    End If
    If Npclist(NpcIndex).MiMana < (Npclist(NpcIndex).mana / 2) Then
        ' Estoy MUY agotado, necesito pociones azules URGENTE
        NMana = 3
    End If
    
    Dim ObjPos As WorldPos
    Dim MasCerca As Integer
    Dim MasDat As String
    
    MasCerca = 100
    
    
    ' Buscamos
    For Y = Npclist(NpcIndex).Pos.Y - MinYBorder + 1 To Npclist(NpcIndex).Pos.Y + MinYBorder - 1
        For X = Npclist(NpcIndex).Pos.X - MinXBorder + 1 To Npclist(NpcIndex).Pos.X + MinXBorder - 1
               If InMapBounds(Map, X, Y) Then
                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                    ObjPos.Map = Map
                    ObjPos.X = X
                    ObjPos.Y = Y
                    ' hay objeto
                    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).TipoPocion = 3 Then
                        ' Cura Vida
                        If NVida >= NMana Then
                            If Distancia(Npclist(NpcIndex).Pos, ObjPos) <= MasCerca Then
                                MasCerca = Distancia(Npclist(NpcIndex).Pos, ObjPos)
                                MasDat = X & "-" & Y
                            End If
                        End If
                    ElseIf ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).TipoPocion = 4 Then
                        ' Cura Mana
                        If NMana >= NVida Then
                            If Distancia(Npclist(NpcIndex).Pos, ObjPos) <= MasCerca Then
                                MasCerca = Distancia(Npclist(NpcIndex).Pos, ObjPos)
                                MasDat = X & "-" & Y
                            End If
                        End If
                    End If
                    
                    'If Len(aNPC_Cosas) > 0 Then Exit Function
                End If
               End If
        Next X
    Next Y
    If MasCerca <> 100 Then
        aNPC_Cosas = MasDat
        Exit Function
    End If
    aNPC_Cosas = ""
End Function

Function aNPC_ALERTA(ByVal NpcIndex As Integer) As Boolean
    Dim Y, X, Map As Integer
    Map = Npclist(NpcIndex).Pos.Map
    aNPC_ALERTA = False
    For Y = Npclist(NpcIndex).Pos.Y - MinYBorder + 1 To Npclist(NpcIndex).Pos.Y + MinYBorder - 1
        For X = Npclist(NpcIndex).Pos.X - MinXBorder + 1 To Npclist(NpcIndex).Pos.X + MinXBorder - 1
               If InMapBounds(Map, X, Y) Then
                    If MapData(Map, X, Y).UserIndex > 0 Then
                        If (UserList(MapData(Map, X, Y).UserIndex).flags.Privilegios < 1 And EsAdmin(MapData(Map, X, Y).UserIndex) = False) Then
                            aNPC_ALERTA = True
                            Exit Function
                        End If
                    End If
               End If
        Next X
    Next Y
End Function

Function aNPC_Amigo(ByVal NpcIndex As Integer) As Integer
    Dim Y, X, Map As Integer
    Map = Npclist(NpcIndex).Pos.Map
    Dim VidaInd As Integer
    Dim VidaMin As Integer
    aNPC_Amigo = 0
    For Y = Npclist(NpcIndex).Pos.Y - MinYBorder + 1 To Npclist(NpcIndex).Pos.Y + MinYBorder - 1
        For X = Npclist(NpcIndex).Pos.X - MinXBorder + 1 To Npclist(NpcIndex).Pos.X + MinXBorder - 1
               If InMapBounds(Map, X, Y) Then
                    If MapData(Map, X, Y).NpcIndex > 0 And MapData(Map, X, Y).NpcIndex <> NpcIndex Then
                        If Npclist(NpcIndex).Name = Npclist(MapData(Map, X, Y).NpcIndex).Name Then
                            If VidaMin > Npclist(MapData(Map, X, Y).NpcIndex).Stats.MinHP Or Npclist(MapData(Map, X, Y).NpcIndex).Meditando = True Then
                                VidaMin = Npclist(MapData(Map, X, Y).NpcIndex).Stats.MinHP
                                VidaInd = MapData(Map, X, Y).NpcIndex
                            End If
                        End If
                    End If
               End If
        Next X
    Next Y
    If VidaMin > 0 Then aNPC_Amigo = VidaInd
End Function

Function aNPC_PathFinding(ByVal NpcIndex As Integer) As Boolean
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
     For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10   '5 tiles in every direction

         'Make sure tile is legal
         If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
             'look for a user
             If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                  'We have to invert the coordinates, this is because
                  'ORE refers to maps in converse way of my pathfinding
                  'routines.
                    ' [GS] No serguir o mirar a a GM's
                    If (UserList(tmpUserIndex).flags.Privilegios > 1 Or EsAdmin(tmpUserIndex)) Or UserList(tmpUserIndex).flags.PocionRepelente = True Then
                        GoTo siguienteNW
                    End If
                    ' [/GS]
                    
                ' [GS] Atacando con magias?
                If UserList(tmpUserIndex).flags.Muerto = 0 Then
                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                        If UserList(tmpUserIndex).Stats.UserAtributos(Suerte) = 0 Or RandomNumber(1, UserList(tmpUserIndex).Stats.UserAtributos(Suerte) + 10) > UserList(tmpUserIndex).Stats.UserAtributos(Suerte) Then
                            Call aNPC_NpcLanzaUnSpell(NpcIndex, tmpUserIndex)
                        End If
                    End If
                End If
                ' [/GS]
                  
                  Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
                  Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
                  Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                  Call SeekPath(NpcIndex)
                  Exit Function
             End If
             
         End If
siguienteNW:
     Next X
 Next Y
End Function

Sub aNPC_NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
Dim i, Spell As Integer

If Npclist(NpcIndex).Movement <> 11 Then Exit Sub
If Npclist(NpcIndex).flags.LanzaSpells <= 0 Then Exit Sub

' [GS] No serguir o mirar a a GM's
If (UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex)) Or UserList(UserIndex).flags.PocionRepelente = True Then Exit Sub
' [/GS]

If UserList(UserIndex).flags.Invisible = 1 And Npclist(NpcIndex).flags.AtacaInvis = True And UserList(UserIndex).flags.TieneMensaje = True Then
' Es un invi evidente ;) y se la prendemos igual
ElseIf UserList(UserIndex).flags.Invisible = 1 Then
    ' Esta invi, pero el npc lo le pega :P
    Exit Sub
End If

' Si el usuario medita, aprobechamos para hacerlo nosotros tambien
If UserList(UserIndex).flags.Meditando = True Then
    If Hechizos(Spell).MiMana < Npclist(NpcIndex).mana Then
        Call NPCMeditando(NpcIndex, True)
    End If
End If

' Busco el hechizo de mayor daño
Dim SpX As Integer  ' daño
Dim SpI As Integer  ' indice
Dim SpM As Integer  ' mana requerido
SpX = 0
SpI = 0
SpM = 0
For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
    If UserList(UserIndex).Stats.MinHP < Hechizos(Npclist(NpcIndex).Spells(i)).MinHP And Hechizos(Npclist(NpcIndex).Spells(i)).SubeHP = 2 Then
        ' Si el user, tiene menos vida que el hechizo, lo elegimos
        SpI = Npclist(NpcIndex).Spells(i)
        Exit For
    Else
        If (Hechizos(Npclist(NpcIndex).Spells(i)).MinHP + 10) > SpX And Hechizos(Npclist(NpcIndex).Spells(i)).SubeHP = 2 Then
            If (Hechizos(Npclist(NpcIndex).Spells(i)).ManaRequerido) < SpM Then
                SpX = Hechizos(Npclist(NpcIndex).Spells(i)).MinHP
                SpI = Npclist(NpcIndex).Spells(i)
                SpM = Hechizos(Npclist(NpcIndex).Spells(i)).ManaRequerido
            End If
        End If
    End If
Next
Spell = SpI

If UserList(UserIndex).flags.Paralizado = 0 Then ' si el user no esta paralizado
    For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
        If Hechizos(Npclist(NpcIndex).Spells(i)).Paraliza = 1 Then Spell = Npclist(NpcIndex).Spells(i)
    Next
End If

'Dim k As Integer
'k = RandomNumber(1, Npclist(NPCindex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Spell)

End Sub
