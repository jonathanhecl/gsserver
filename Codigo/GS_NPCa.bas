Attribute VB_Name = "GS_NPCa"
'
' ###############
' ## NPCa v0.2 ##
' ###############
'  NPC de Agite
'
' Moficiado:
'   10/03/2005 (Se inicia la v0.2)
'       NPCa_EstaPegado - Termiando
'       NPCa_ObjetivoUser - Terminado
'       NPCa_TengoItem - Terminado
'       NPCa_Cerebro - Empezado
'
' Objetivos:
' * Crear un NPC que tenga movimiento como usuarios
' * Que se protejan entre otros NPCs de su misma raza
' ##################################################################
' * Que tomen objetos y los utilizen inteligentemente (PROXIMAMENTE)
' ##################################################################
' * Que utilize magias, de forma inteligente
' * Que obtenga una personalidad unica
'
' Programado por ^[GS]^
'
Type NPCa_UserStat
    Userindex As Integer
    Vida As Integer
    Nivel As Integer
    Meditando As Boolean
End Type

' Proceso Principal
Sub NPCa_Cerebro(ByVal NPCindex As Integer)
' Criterios:
' * Si un user me ataco...
'    ...ataco (golpe/magia-50%)
' ##################################################################
'    ...si hay objetos en el piso (que nos recuperaria) los agarro
' ##################################################################
'    ...me recupero (curo/medito-50%)
' * Si tengo user en la mira, pero no me ha atacado...
'    ...reviso si un "amigo" esta siendo atacado por el, si es asi...
'       ...curo a mi amigo (magia)
'       ...ataco a el usuario (golpe/magia)
'       ...me recupeto (curo/medito)
' * Si no hay users, pero si "amigos"...
' ##################################################################
'   ...si hay objetos en el piso (que nos recuperaria) los agarro
' ##################################################################
'   ...si estoy herido me curo
'   ...si tengo un amigo herido lo curo
'   ...nos juntamos todos los "amigos" y esperamos
' * Si estoy solo
' ##################################################################
'   ...si hay objetos en el piso (que nos recuperaria) los agarro
' ##################################################################
'   ...si estoy herido me curo
'   ...espero

' Si un user me ataco...
If Npclist(NPCindex).flags.AttackedIndex <> 0 Then
    If UserList(Npclist(NPCindex).flags.AttackedIndex).flags.Muerto = False Then
        ' ...si no esta muerto...
        
        ' ##################################################################
        'If (RandomNumber(1, 100)) < 22 And NPCa_HayObjUTIL(NPCindex) Then
            ' Busco objtetos UTILES (22%)
        'End If
        ' ##################################################################
        
        'If (RandomNumber(1, 100)) < 22 And NPCa_NCYP(NPCindex) Then
            ' Me recupero (22%)
        'End If
        
        ' ...ataco (50%)
        If NPCa_EstaPegado(NPCindex, Npclist(NPCindex).flags.AttackedIndex) Then
            If Npclist(NPCindex).flags.LanzaSpells > 0 And (RandomNumber(1, 100)) > 50 Then
                ' Tiene hechizos (con un 50 % de probabilidades)
                Call NPCa_LanzarHechizos(NPCindex, Npclist(NPCindex).flags.AttackedIndex)
                Exit Sub
            End If
            ' Ataco con golpe
            Dim tHeading As Byte
            tHeading = FindDirection(Npclist(NPCindex).Pos, UserList(Npclist(NPCindex).flags.AttackedIndex).Pos)
            Call MoveNPCChar(NPCindex, tHeading)
            Call NpcAtacaUser(NPCindex, Npclist(NPCindex).flags.AttackedIndex)
            ' ##################################################################
            ' Hago un paso hacia atras
            ' Call NPCa_Esquivar
            ' ##################################################################
        Else    ' No esta cerca (voy y utilizo magias)
            ' Voy hacia el la posicion del Usuario
            ' Call NPCa_IrA(NPCindex, UserList(Npclist(NPCindex).flags.AttackedIndex).Pos)
            ' Si tengo magia, la uso
            Call NPCa_LanzarHechizos(NPCindex, Npclist(NPCindex).flags.AttackedIndex)
        End If
        
    End If
End If

' Lo que falta:
' ##################################################################
' NPCa_HayObjUTIL   - Buscar, si hay objetos Utiles en el suelo
' NPCa_Esquivar     - Esquivar/Alejarse
' ##################################################################
' NPCa_NCYP         - Necesito Curarme y Puedo?
' NPCa_IrA          - Ir a X,Y

End Sub



' Revisar si el User esta pegado al NPC
Function NPCa_EstaPegado(ByVal NPCindex As Integer, ByVal Userindex As Integer) As Boolean
' Devuelve Verdadero si Hay un usuario pegado al NPC
NPCa_EstaPegado = False
If Npclist(NPCindex).Pos.x = UserList(Userindex).Pos.x + 1 And Npclist(NPCindex).Pos.y = UserList(Userindex).Pos.y Then
    NPCa_EstaPegado = True
ElseIf Npclist(NPCindex).Pos.x = UserList(Userindex).Pos.x - 1 And Npclist(NPCindex).Pos.y = UserList(Userindex).Pos.y Then
    NPCa_EstaPegado = True
ElseIf Npclist(NPCindex).Pos.y = UserList(Userindex).Pos.y - 1 And Npclist(NPCindex).Pos.x = UserList(Userindex).Pos.x Then
    NPCa_EstaPegado = True
ElseIf Npclist(NPCindex).Pos.y = UserList(Userindex).Pos.y + 1 And Npclist(NPCindex).Pos.x = UserList(Userindex).Pos.x Then
    NPCa_EstaPegado = True
End If
End Function

' Revisa si el NPC tiene Item en el inventario
Function NPCa_TengoItem(ByVal NPCindex As Integer, ByVal OBJindex As Integer) As Integer
' Devuelve en que posicion del inventario se encuentra
NPCa_TengoItem = 0
Dim Slot As Integer
For Slot = 1 To MAX_INVENTORY_SLOTS
    If Npclist(NPCindex).Invent.Object(Slot).OBJindex = OBJindex Then
        NPCa_TengoItem = Slot
        Exit Function
    End If
Next
End Function

' Busca en el area de vision, si hay algun usuario, y atacar al que presente mayor amenaza
Function NPCa_ObjetivoUser(ByVal NPCindex As Integer) As Integer
' Devuelve el Userindex
NPCa_ObjetivoUser = 0
Dim y, x, Map, Userindex As Integer
Dim Debil As NPCa_UserStat
Debil.Vida = STAT_MAXHP
Debil.Nivel = MaxNivel
Debil.Meditando = False
Debil.Userindex = 0

' Criterios:
' * Si me esta atacando, es mi enemigo
' * Si tiene menor vida que todos, entonces ataco a este...
'   ...Si tiene la misma vida...
'       ...ataco al que esta Meditando
'       ...ataco al que tiene menor Nivel

Map = Npclist(NPCindex).Pos.Map
For y = Npclist(NPCindex).Pos.y - MinYBorder + 1 To Npclist(NPCindex).Pos.y + MinYBorder - 1
    For x = Npclist(NPCindex).Pos.x - MinXBorder + 1 To Npclist(NPCindex).Pos.x + MinXBorder - 1
            If InMapBounds(Map, x, y) Then
                If MapData(Map, x, y).Userindex > 0 Then
                    Userindex = MapData(Map, x, y).Userindex
                    If (UserList(Userindex).flags.Privilegios < 1 And EsAdmin(Userindex) = False) Then
                        ' Si hay un usuario.....
                        If UserList(Userindex).flags.Muerto <> 0 Then
                            ' Si esta vivo...
                            If Npclist(NPCindex).flags.AttackedIndex = Userindex Then
                                ' Me esta atacando
                                NPCa_ObjetivoUser = Userindex
                                Exit Function
                            Else
                                If UserList(Userindex).Stats.MinHP < Debil.Vida Then
                                    ' Tiene menos vida que el anterior
                                    Debil.Vida = UserList(Userindex).Stats.MinHP
                                    Debil.Nivel = UserList(Userindex).Stats.ELV
                                    Debil.Meditando = UserList(Userindex).flags.Meditando
                                    Debil.Userindex = Userindex
                                ElseIf UserList(Userindex).Stats.MinHP = Debil.Vida Then
                                    ' Tiene la misma vida que el anterior
                                    If Debil.Meditando = False And UserList(Userindex).flags.Meditando = True Then
                                        ' Si el objetivo anterior no estaba meditando, y este si, entonces este es mejor objetivo.
                                        Debil.Vida = UserList(Userindex).Stats.MinHP
                                        Debil.Nivel = UserList(Userindex).Stats.ELV
                                        Debil.Meditando = UserList(Userindex).flags.Meditando
                                        Debil.Userindex = Userindex
                                    ElseIf UserList(Userindex).Stats.ELV < Debil.Nivel Then
                                        ' Si este user tiene menor nivel, que el otro, pegarle a este
                                        Debil.Vida = UserList(Userindex).Stats.MinHP
                                        Debil.Nivel = UserList(Userindex).Stats.ELV
                                        Debil.Meditando = UserList(Userindex).flags.Meditando
                                        Debil.Userindex = Userindex
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    Next x
Next y
' Encontre un bueno objetivo?
If Debil.Userindex <> 0 Then
    ' Obtengo a este objetivo, sin que me atacaran previamente
    NPCa_ObjetivoUser = Debil.Userindex
End If
End Function
