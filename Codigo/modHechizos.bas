Attribute VB_Name = "modHechizos"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit




Function ModMagic(ByVal UserIndex As Integer, ByVal Damange As Double) As Double
On Error Resume Next
ModMagic = Damange
' [GS] Arma que modifica magia
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic) = True Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic & FONTTYPE_INFX)
            ModMagic = ModMagic * ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic
        End If
    End If
End If
' [GS] Accesorios que modifican magia
If UserList(UserIndex).Invent.Accesorio1EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic & FONTTYPE_INFX)
            ModMagic = ModMagic * ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic
        End If
    End If
End If
If UserList(UserIndex).Invent.Accesorio2EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic & FONTTYPE_INFX)
            ModMagic = ModMagic * ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic
        End If
    End If
End If
End Function


Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
Dim i As Byte
If Npclist(NpcIndex).Meditando = True Then Exit Sub
If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then Exit Sub

' [GS] Hay consulta?
If HayConsulta = True Then
        If (UserList(QuienConsulta).Pos.Map = Npclist(NpcIndex).Pos.Map) Then     ' NPC?
            If Distancia(Npclist(NpcIndex).Pos, UserList(QuienConsulta).Pos) < 18 Or Distancia(UserList(UserIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then Exit Sub
        End If
End If
' [/GS]
' [GS] Corrigue el bug del NPC que ataca por portales
If Npclist(NpcIndex).Pos.Map <> UserList(UserIndex).Pos.Map Then Exit Sub
' [/GS]
If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 And Npclist(NpcIndex).flags.AtacaInvis = True And UserList(UserIndex).flags.TieneMensaje = True Then
' Es un invi evidente ;) y se la prendemos igual
ElseIf UserList(UserIndex).flags.Invisible = 1 Then
    ' Esta invi, pero el npc lo le pega :P
    Exit Sub
End If

Call ToBienSpell(NpcIndex, Spell)
If Npclist(NpcIndex).Meditando = True Then Exit Sub

' No tirar remover al enemigo
If Hechizos(Spell).RemoverParalisis = 1 Then
    For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
        If Hechizos(i).MinHP > 0 Then Spell = Npclist(NpcIndex).Spells(i)
    Next
End If

Call ToBienSpell(NpcIndex, Spell)
If Npclist(NpcIndex).Meditando = True Then Exit Sub

' 0.12b3
If RandomNumber(0, 4) > 2 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Long

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
        
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    ' /* v0.12a12
    
    'accesorios 1
    If (UserList(UserIndex).Invent.Accesorio1EqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).DefensaMagicaMax + 1)
    End If
    'accesorios 2
    If (UserList(UserIndex).Invent.Accesorio2EqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).DefensaMagicaMax + 1)
    End If
    ' armaduras
    If (UserList(UserIndex).Invent.ArmourEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax + 1)
    End If
    
    ' v0.12a12 */

    
    If daño < 0 Then daño = 0
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    If (UserList(UserIndex).flags.Privilegios = 0 And EsAdmin(UserIndex) = False) Then
        If UserList(UserIndex).flags.PocionRepelente = False Then
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
        End If
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(UserIndex).Stats.MinHP < 1 Then
        UserList(UserIndex).Stats.MinHP = 0
        Call UserDie(UserIndex)
    End If
    
End If

' ### ACTUALIZA ESTADO ###
Call SendUserStatsBox(val(UserIndex))
' ### ACTUALIZA ESTADO ###

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then ' Tiene Ropa
            If CInt(RandomNumber(1, 100)) <= ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).NoParalisis And ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).NoParalisis > 0 Then
                ' No paraliza
                Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha intentado paralizar." & FONTTYPE_FIGHT)
            Else
                UserList(UserIndex).flags.Paralizado = 1
                QuitarLAGalUser (UserIndex) ' ### QUITA EL LAG ###
                UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
                Call SendData(ToIndex, UserIndex, 0, "PARADOK")
            End If
        Else
                UserList(UserIndex).flags.Paralizado = 1
                QuitarLAGalUser (UserIndex) ' ### QUITA EL LAG ###
                UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
                Call SendData(ToIndex, UserIndex, 0, "PARADOK")
        End If
     End If
End If

If Npclist(NpcIndex).TieneMana = True Then
    Npclist(NpcIndex).MiMana = Npclist(NpcIndex).MiMana - Hechizos(Spell).ManaRequerido
    If Npclist(NpcIndex).MiMana < 0 Then Npclist(NpcIndex).MiMana = 0
    'Call SendData(ToIndex, Userindex, 0, "||Mana: " & Npclist(NPCindex).MiMana & FONTTYPE_INFX)
    If (Npclist(NpcIndex).MiMana < Npclist(NpcIndex).mana) Then Call NPCMeditando(NpcIndex, True)
End If

Call SendData(ToPCArea, UserIndex, Npclist(NpcIndex).Pos.Map, "||" & vbCyan & "°" & Hechizos(Spell).PalabrasMagicas & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
Npclist(NpcIndex).flags.Hablo = True

End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "CFX" & Npclist(TargetNPC).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        ' [GS] Habla
        Call SendData(ToPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "||" & vbCyan & "°" & Hechizos(Spell).PalabrasMagicas & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
        Npclist(NpcIndex).flags.Hablo = True
        ' [/GS]
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        ' ### HECHI CLASE ###
        If Hechizos(hIndex).ExclusivoClase <> 0 Then ' que tiene algo
            If UserList(UserIndex).clase <> Hechizos(hIndex).ExclusivoClase Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes aprender hechizos que no son para tu clase." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        ' ### HECHI CLASE ###
        ' [GS] Hechizo de requisitos
        If Hechizos(hIndex).Requiere > 0 Then
            Dim LoPuede As Boolean
            Dim SlotX As Integer
            LoPuede = False
            For SlotX = 1 To 20
                If UserList(UserIndex).Invent.Object(SlotX).Equipped = 1 Then
                    ' esta equipado
                    If UserList(UserIndex).Invent.Object(SlotX).ObjIndex = Hechizos(hIndex).Requiere Then
                        LoPuede = True
                        Exit For
                    End If
                End If
            Next
            If LoPuede = False Then
                If Hechizos(hIndex).Requiere < NumObjDatas Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes " & ObjData(Hechizos(hIndex).Requiere).Name & " equipado/a." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        ' [/GS]
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal UserIndex As Integer)
On Error Resume Next
    If S = "" Then Exit Sub
    Dim ind As String
    ind = UserList(UserIndex).Char.CharIndex
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbCyan & "°" & S & "°" & ind)
    UserList(UserIndex).flags.TieneMensaje = True
    Exit Sub
End Sub
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
' ### HECHI CLASE ###
If Hechizos(HechizoIndex).ExclusivoClase <> 0 Then ' que tiene algo
    If (UserList(UserIndex).clase) <> (Hechizos(HechizoIndex).ExclusivoClase) Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes lanzar hechizos que no son para tu clase." & FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
End If
' ### HECHI CLASE ###

If UserList(UserIndex).flags.TiraExp = True And Hechizos(HechizoIndex).Tipo = uExplocionMagica Then
    Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
    UserList(UserIndex).flags.TiraExp = False
    PuedeLanzar = False
    Exit Function
End If

' [GS] Esta tirando una explocion magica??
If UserList(UserIndex).flags.TiraExp = True Then
    Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
    UserList(UserIndex).flags.TiraExp = False
End If
' [/GS]

' [GS] Hechizo de requisitos
If Hechizos(HechizoIndex).Requiere > 0 Then
    Dim LoPuede As Boolean
    Dim SlotX As Integer
    LoPuede = False
    For SlotX = 1 To 20
        If UserList(UserIndex).Invent.Object(SlotX).Equipped = 1 Then
            ' esta equipado
            If UserList(UserIndex).Invent.Object(SlotX).ObjIndex = Hechizos(HechizoIndex).Requiere Then
                LoPuede = True
                Exit For
            End If
        End If
    Next
    If LoPuede = False Then
        If Hechizos(HechizoIndex).Requiere < NumObjDatas Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & ObjData(Hechizos(HechizoIndex).Requiere).Name & " equipado/a." & FONTTYPE_INFO)
            Exit Function
        End If
    End If
End If
' [/GS]
' [GS] Esta en torneo!
If UserList(UserIndex).Pos.Map = MapaDeTorneo And HayTorneo = True Then
    If ConfigTorneo > 1 Then
        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
                ' no vale INVI!
                Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar Invisibilidad en este Torneo!!" & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
        End If
        If Hechizos(HechizoIndex).Ceguera = 1 Then
               ' no vale CEGUERA!
               Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar Ceguera en este Torneo!!" & FONTTYPE_INFO)
               PuedeLanzar = False
               Exit Function
        End If
        If Hechizos(HechizoIndex).Estupidez = 1 Then
                ' no vale ESTUPIDEZ!
                Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar Estupidez en este Torneo!!" & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
        End If
        If Hechizos(HechizoIndex).Invoca > 0 Then
                ' no valen MASCOTAS
                Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar Mascotas en este Torneo!!" & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
        End If
        If ConfigTorneo = 3 Then
            If Hechizos(HechizoIndex).Paraliza = 1 Then
                ' no vale paralizar
                Call SendData(ToIndex, UserIndex, 0, "||No puedes Paralizar a tu oponente en este Torneo!!" & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    If ConfigTorneo = 4 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes utilizar hechizos en este Torneo!!" & FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    End If
End If
' [/GS]

' [GS] Modo consulta??
'If HayConsulta = True Then
'        If UserList(UserIndex).flags.Privilegios < 1 And (UserList(QuienConsulta).Pos.Map = UserList(UserIndex).Pos.Map) Then ' User?
'            If Distancia(UserList(UserIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then
'                Call SendData(ToIndex, UserIndex, 0, "||No puedes utilizar hechizos en medio de una consulta!!" & FONTTYPE_INFO)
'                Exit Function
'            End If
'        End If
'End If
' [/GS]
If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        ' 0.12b3
        If UCase$(UserList(UserIndex).clase) = CLASS_MAGO Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(ToIndex, UserIndex, 0, "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar este conjuro sin la ayuda de un báculo." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
    
    If Distancia(UserList(UserIndex).Pos, wp2) > 18 Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & "INTENDO HACK!!! IP:" & UserList(UserIndex).ip & " trato de lanzar un spell desde mucha distancia!.")
            'Call Cerrar_Usuario(UserIndex)
            Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Estás muy cansado para lanzar este hechizo." & FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes puntos de magia para lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call SendData(ToIndex, UserIndex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If
End Function


Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        
                        ' Anti-Radar???
                        ' 0.12b1
                        'Call EraseUserChar(ToIndex, MapData(PosCasteadaM, TempX, TempY).UserIndex, 0, MapData(PosCasteadaM, TempX, TempY).UserIndex)
                        'Call MakeUserChar(ToMap, 0, UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Pos.Map, MapData(PosCasteadaM, TempX, TempY).UserIndex, PosCasteadaM, TempX, TempY)
                        'Call SendData(ToMap, 0, UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Pos.Map, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
                        
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
                        End If
                        
                        ' Anti-Radar???
                        ' 0.12b1
                        'Call EraseUserChar(ToMap, MapData(PosCasteadaM, TempX, TempY).UserIndex, 0, MapData(PosCasteadaM, TempX, TempY).UserIndex)
                        'Call MakeUserChar(ToIndex, MapData(PosCasteadaM, TempX, TempY).UserIndex, UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Pos.Map, MapData(PosCasteadaM, TempX, TempY).UserIndex, PosCasteadaM, TempX, TempY)
                        'Call SendData(ToIndex, MapData(PosCasteadaM, TempX, TempY).UserIndex, UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Pos.Map, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
                        
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If

End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

'Call LogTarea("HechizoInvocacion")

' [GS] Modo Counter?
If UserList(UserIndex).flags.CS_Esta = True And UserList(UserIndex).Pos.Map = MapaCounter Then
    Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar en este mapa!" & "~255~255~0~1~0")
    Exit Sub
End If
' [/GS]

' [GS] Torneos
If HayTorneo = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    If MaxMascotasTorneo = 0 Then
    ' No se puede
        Call SendData(ToIndex, UserIndex, 0, "||No se pueden invocar mascotas en este torneo." & FONTTYPE_INFO)
        Exit Sub
    ElseIf UserList(UserIndex).NroMacotas >= MaxMascotasTorneo Then
    ' Lego al maximo
        Call SendData(ToIndex, UserIndex, 0, "||Has llegado al maximo permitido de mascotas para este torneo." & FONTTYPE_INFO)
        Exit Sub
    End If
End If
' [/GS]

If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

Dim h As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(h).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
        If ind <= MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uInvocacion '
       Call HechizoInvocacion(UserIndex, b)
    ' [GS] Explocion Magica
    Case uExplocionMagica
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
            UserList(UserIndex).flags.NumHechExp = uh
            UserList(UserIndex).flags.TimerExp = 0
            UserList(UserIndex).flags.XExp = UserList(UserIndex).flags.TargetX
            UserList(UserIndex).flags.YExp = UserList(UserIndex).flags.TargetY
            UserList(UserIndex).flags.TiraExp = True
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Estas en una zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
    ' [/GS]
    Case uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
    

End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)

If AntiLukers Then
    ' [GS] Anti robo de NPC
    ' Este NPC fue atacado por alguien?
    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex > 0 Then
        ' Su atacante aun esta cerca Y no soy Yo
        If UserList(Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex).flags.SuNPC = UserList(UserIndex).flags.TargetNPC And UserIndex <> Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex Then
            ' No es el dueño
            If Criminal(Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex) = False Then
                If UserList(UserIndex).flags.Seguro = True Then
                    Call SendData(ToIndex, UserIndex, 0, "||Este NPC ya tiene un dueño, para atacarlo igual debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT_YO)
                    Exit Sub
                ElseIf UserList(UserIndex).flags.Seguro = False Then
                    ' Se vuelve crimi
                    Call xRobar(UserIndex)
                    Call SendData(ToIndex, Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex, 0, "||" & UserList(UserIndex).Name & " te ha robado a " & Npclist(UserList(UserIndex).flags.TargetNPC).Name & FONTTYPE_FIGHT)
                    Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex = UserIndex
                    ' Ataka
                End If
            Else
                Call SendData(ToIndex, Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado a " & Npclist(UserList(UserIndex).flags.TargetNPC).Name & FONTTYPE_FIGHT)
                Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex = UserIndex
                ' Ataka
            End If
        Else
            ' Ataka
        End If
    Else
        Npclist(UserList(UserIndex).flags.TargetNPC).flags.AttackedIndex = UserIndex
        ' Ataka
    End If
    ' [/GS]
End If

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

    Call SendUserStatsBox(UserIndex)
End If

End Sub
Sub LanzarHechizo(Index As Integer, UserIndex As Integer)



Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case uUsuarios ' Hechizo para usar sobre Usuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case uNPC ' Hechizo para usar sobre NPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case uUsuariosYnpc ' Hechizo para usar sobre Usuarios y NPCs
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
        Case uTerreno ' Hechizo para usar sobre el Terreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If
                

End Sub
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)


Dim h As Integer, TU As Integer
h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser

If Hechizos(h).Invisibilidad = 1 Then
   UserList(TU).flags.Invisible = 1
   
   ' Anti-Radar?? :S
   ' v0.12b1
   'Call EraseUserChar(ToMap, 0, UserList(TU).Pos.Map, TU)
   'Call MakeUserChar(ToIndex, TU, 0, TU, UserList(TU).Pos.Map, UserList(TU).Pos.X, UserList(TU).Pos.Y)
   'Call SendData(ToIndex, TU, 0, "NOVER" & UserList(TU).Char.CharIndex & ",1")
   
   Call SendData(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
   Call InfoHechizo(UserIndex)
   b = True
End If

If Hechizos(h).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios > 0 Or EsAdmin(TU) = True Then
        Exit Sub
    End If
       
   If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto." & FONTTYPE_INFO)
        Exit Sub
   End If
    
    'copio el char original al mimetizado
    
    UserList(UserIndex).CharMimetizado.Body = UserList(UserIndex).Char.Body
    UserList(UserIndex).CharMimetizado.Head = UserList(UserIndex).Char.Head
    UserList(UserIndex).CharMimetizado.CascoAnim = UserList(UserIndex).Char.CascoAnim
    UserList(UserIndex).CharMimetizado.ShieldAnim = UserList(UserIndex).Char.ShieldAnim
    UserList(UserIndex).CharMimetizado.WeaponAnim = UserList(UserIndex).Char.WeaponAnim
    
   UserList(UserIndex).flags.Mimetizado = 1
   
    'ahora pongo local el del enemigo
    UserList(UserIndex).Char.Body = UserList(TU).Char.Body
    UserList(UserIndex).Char.Head = UserList(TU).Char.Head
    UserList(UserIndex).Char.CascoAnim = UserList(TU).Char.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(TU).Char.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(TU).Char.WeaponAnim
    
    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "CP" & UserList(UserIndex).Char.CharIndex & "," & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & UserList(UserIndex).Char.CascoAnim)
   
   Call InfoHechizo(UserIndex)
   b = True

End If


If Hechizos(h).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
End If

If Hechizos(h).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(h).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
End If

If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(h).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            If UserList(TU).Invent.ArmourEqpObjIndex > 0 Then ' Tiene Ropa
                If CInt(RandomNumber(1, 100)) <= ObjData(UserList(TU).Invent.ArmourEqpObjIndex).NoParalisis And ObjData(UserList(TU).Invent.ArmourEqpObjIndex).NoParalisis > 0 Then
                    ' No paraliza
                    Call SendData(ToIndex, TU, 0, "||" & UserList(UserIndex).Name & " te ha intentado paralizar." & FONTTYPE_FIGHT)
                Else
                    UserList(TU).flags.Paralizado = 1
                    Call QuitarLAGalUser(TU)  ' ### QUITA EL LAG ###
                    UserList(TU).Counters.Paralisis = IntervaloParalizado
                    Call SendData(ToIndex, TU, 0, "PARADOK")
                    Call InfoHechizo(UserIndex)
                    b = True
                    If UserIndex <> TU Then
                        Call UsuarioAtacadoPorUsuario(UserIndex, TU)
                    End If
                End If
            Else
                UserList(TU).flags.Paralizado = 1
                Call QuitarLAGalUser(TU)  ' ### QUITA EL LAG ###
                UserList(TU).Counters.Paralisis = IntervaloParalizado
                Call SendData(ToIndex, TU, 0, "PARADOK")
                Call InfoHechizo(UserIndex)
                b = True
                If UserIndex <> TU Then
                    Call UsuarioAtacadoPorUsuario(UserIndex, TU)
                End If
            End If
    End If
End If

If Hechizos(h).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
                UserList(TU).flags.Paralizado = 0
                Call SendData(ToIndex, TU, 0, "PARADOK")
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If
' [GS]
If Hechizos(h).RemoverCeguera = 1 Then
    If UserList(TU).flags.Ceguera = 1 Then
        UserList(TU).flags.Ceguera = 0
        UserList(TU).Counters.Ceguera = 0
        Call SendData(ToIndex, UserIndex, 0, "NSEGUE")
        Call InfoHechizo(UserIndex)
    End If
End If
If Hechizos(h).RemoverEstupidez = 1 Then
    If UserList(TU).flags.Estupidez = 1 Then
        UserList(TU).flags.Ceguera = 0
        UserList(TU).Counters.Ceguera = 0
        Call SendData(ToIndex, UserIndex, 0, "NESTUP")
        Call InfoHechizo(UserIndex)
    End If
End If

' [/GS]

If Hechizos(h).Revivir = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        If Not Criminal(TU) Then
                If TU <> UserIndex Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                    Call SendData(ToIndex, UserIndex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
                End If
        End If
        
        Call RevivirUsuario(TU)
    End If
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(h).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        b = True
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
End If

If Hechizos(h).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "DUMB")
        Call InfoHechizo(UserIndex)
        b = True
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
End If

' [GS] AutoComentarista
If HayTorneo = True And AutoComentarista = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    If TU <> UserIndex Then
        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " tira " & Hechizos(h).nombre & " sobre " & UserList(UserIndex).Name & FONTTYPE_INFO)
    Else
        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " se tira " & Hechizos(h).nombre & FONTTYPE_INFO)
    End If
End If
' [/GS]

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

' [GS] No magias?
If Npclist(NpcIndex).NoMagias = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a este NPC con hechizos." & FONTTYPE_INFO)
    Exit Sub
End If
' [/GS]

If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_FIGHT_YO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   Call NpcAtacado(NpcIndex, UserIndex)
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_FIGHT_YO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 1
   b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

' 0.12b3
If Hechizos(hIndex).Inmoviliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
   End If
End If

If Hechizos(hIndex).Paraliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).flags.Inmovilizado = 0
            'QuitarLAGalUser (UserIndex) ' ### QUITA EL LAG ###
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El NPC es inmune a este hechizo." & FONTTYPE_FIGHT_YO)
   End If
   Call NpcAtacado(NpcIndex, UserIndex)
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||No puedes remover a este NPC." & FONTTYPE_FIGHT_YO)
   End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
On Error GoTo Errores
Dim Calculo As Long ' Hiper-AO
Dim daño As Double

' [GS] Hay consulta?
'If HayConsulta = True Then
'        If UserList(UserIndex).flags.Privilegios < 1 And (UserList(QuienConsulta).Pos.Map = Npclist(NpcIndex).Pos.Map) Then    ' NPC?
'            If Distancia(Npclist(NpcIndex).Pos, UserList(QuienConsulta).Pos) > 18 Or Distancia(UserList(UserIndex).Pos, UserList(QuienConsulta).Pos) > 18 Then
'                Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar hechizos en una consulta." & FONTTYPE_FIGHT)
'                Exit Sub
'            End If
'        End If
'End If
' [/GS]

' [GS] No magias?
If Npclist(NpcIndex).NoMagias = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a este NPC con hechizos." & FONTTYPE_INFO)
    Exit Sub
End If
' [/GS]

' [GS] No ataca guiardias con seguro on
If UserList(UserIndex).flags.Seguro = True And Npclist(NpcIndex).TargetNPC = NPCTYPE_GUARDIAS Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes a un Guardia con el seguro activado." & FONTTYPE_INFO)
    Exit Sub
End If
' [/GS]
daño = 0
'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = ModMagic(UserIndex, daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV))
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, daño, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, UserIndex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT_YO)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Npclist(NpcIndex).MaestroUser <= 0 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese NPC." & FONTTYPE_FIGHT_YO)
            Exit Sub
        End If
    Else
        If MapInfo(Npclist(NpcIndex).Pos.Map).Pk = False Then
            Call SendData(ToIndex, UserIndex, 0, "||No podés atacar mascotas en zonas seguras" & FONTTYPE_FIGHT_YO)
            Exit Sub
        End If
    End If
    
    ' Toma el daño del hechizo!!
    If Hechizos(hIndex).MinHP < Hechizos(hIndex).MaxHP Then
        daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Else
        daño = Hechizos(hIndex).MaxHP
    End If
    
    ' Subimos el bonus extra por el Nivel
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    daño = CLng(ModMagic(UserIndex, daño))
    
    ' 0.12b3
    If UCase$(UserList(UserIndex).clase) = CLASS_MAGO Then
        If Hechizos(hIndex).StaffAffected Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 80% del original
            End If
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    ' [NEW] Hiper-AO
    
    ' [GS] Barrera?
    If Npclist(NpcIndex).flags.BarreraEspejo > 0 Then
        If CInt(RandomNumber(0, 100)) <= Npclist(NpcIndex).flags.BarreraEspejo Then
            ' [GS] No serguir o mirar a a GM's
            If (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) And Npclist(NpcIndex).flags.Paralizado = 0 Then
            ' [/GS]
                Call SendData(ToIndex, UserIndex, 0, "||La creatura ha esquivado tu ataque y esta furiosa." & FONTTYPE_WARNING)
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                Dim k As Integer
                If Npclist(NpcIndex).flags.LanzaSpells >= 1 Then
                    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
                End If
                Exit Sub
            ElseIf (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) And Npclist(NpcIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||La creatura se a removido furiosa." & FONTTYPE_WARNING)
                Npclist(NpcIndex).flags.Paralizado = 0
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                Exit Sub
            End If
        End If
    End If
    ' [/GS]
    
    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)
    
    Calculo = (daño / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)

    If daño >= Npclist(NpcIndex).Stats.MinHP Then
    '    If daño >= Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
        Calculo = (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)
    '    Else
    '    Calculo = MiNPC.GiveEXP / 2 + (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP / 2)
    '    End If
    End If

    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    Call SendData(ToIndex, UserIndex, 0, "U2" & daño)
    
'    If UserList(Userindex).flags.Party > 0 And UserList(UserList(Userindex).flags.Party).flags.Muerto = 0 Then ' esta en party?
'        Call AddtoVar(UserList(Userindex).Stats.Exp, CInt(Calculo / 2), MaxExp)
'        Call SendData(ToIndex, Userindex, 0, "||Has ganado " & CInt(Calculo / 2) & " puntos de experiencia." & FONTTYPE_FIGHT)
'        Call AddtoVar(UserList(UserList(Userindex).flags.Party).Stats.Exp, CInt(Calculo / 2), MaxExp)
'        Call SendData(ToIndex, UserList(Userindex).flags.Party, 0, "||Has ganado " & CInt(Calculo / 2) & " puntos de experiencia." & FONTTYPE_FIGHT)
'        Call CheckUserLevel(UserList(Userindex).flags.Party)
'    Else

        Call GanarExp(UserIndex, Calculo, False)

    

    'Controla el nivel del usuario
    
    '' [/NEW]
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Calculo = 0
        Call MuereNpc(NpcIndex, UserIndex)
    Else
        Call CheckPets(NpcIndex, UserIndex)
    End If
End If


Exit Sub

Errores:
    Call LogError("Error en HechizoPropNPC - Err " & Err.Number & " - " & Err.Description)

End Sub
Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim h As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(h).WAV)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(ToPCArea, UserList(UserIndex).flags.TargetUser, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(ToNPCArea, UserList(UserIndex).flags.TargetNPC, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Map, "CFX" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT_YO)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(h).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).PropioMsg & FONTTYPE_FIGHT_YO)
        End If
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " la criatura." & FONTTYPE_FIGHT_YO)
    End If
    
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
On Error GoTo Errores
Dim h As Integer
Dim daño As Long
Dim tempChr As Integer
Dim Miron As String

Miron = "Check Sistema de Consulta"
' [GS] Hay consulta?
If HayConsulta = True Then
        If (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) And (UserList(QuienConsulta).Pos.Map = UserList(UserIndex).Pos.Map) Then    ' User?
            If Distancia(UserList(UserIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes lanzar hechizos en una consulta." & FONTTYPE_FIGHT_YO)
                Exit Sub
            End If
        End If
End If
' [/GS]
    
Miron = "Check Targets"
h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser
      
'Hambre
If Hechizos(h).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHam, _
         daño, UserList(tempChr).Stats.MaxHam)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT_YO)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(h).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

'Sed
If Hechizos(h).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, daño, _
         UserList(tempChr).Stats.MaxAGU)
         
    If UserIndex <> tempChr Then
      Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT_YO)
    End If
    
    b = True
    
ElseIf Hechizos(h).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

' <-------- Fuerza ---------->
If Hechizos(h).SubeFuerza = 1 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If

    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

'Salud
If Hechizos(h).SubeHP = 1 Then

    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If

    Miron = "Curacion"
    daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    daño = ModMagic(UserIndex, daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV))
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, daño, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT_YO)
    End If
    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then
    Miron = "Daño"
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT_YO)
        Exit Sub
    End If
    
    ' [GS]
    If UserList(UserIndex).Pos.Map = MapaAgite Then
        If RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP) > (Hechizos(h).MinHP + Hechizos(h).MaxHP) / 2 Then
            daño = RandomNumber(Hechizos(h).MinHP + 50, Hechizos(h).MaxHP + 50)
        Else
            daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
        End If
    Else
        daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    End If
    ' [/GS]
    
    ' 0.12b4
    daño = CLng(ModMagic(UserIndex, daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)))
    
    ' 0.12b3
    If UCase$(UserList(UserIndex).clase) = CLASS_MAGO Then
        If Hechizos(h).StaffAffected Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                    daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax + 1)
    End If
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax + 1)
    End If
    
    ' /* v0.12a12
    
    'accesorios 1
    If (UserList(tempChr).Invent.Accesorio1EqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.Accesorio1EqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.Accesorio1EqpObjIndex).DefensaMagicaMax + 1)
    End If
    'accesorios 2
    If (UserList(tempChr).Invent.Accesorio2EqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.Accesorio2EqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.Accesorio2EqpObjIndex).DefensaMagicaMax + 1)
    End If
    ' armaduras
    If (UserList(tempChr).Invent.ArmourEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax + 1)
    End If
    
    ' v0.12a12 */
    
    If daño < 0 Then daño = 0
    
    Miron = "Check Armaduras que devuelven ataques"
    ' [GS] Devuelve?
    If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        ' Si tiene una armadura equipada
        If ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).Devuelve > 0 Then
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).Devuelve)
            Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tempChr).Name & " te ha devuelto " & Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).Devuelve) & " puntos de vida." & FONTTYPE_FIGHT)
            daño = daño - Porcentaje(daño, (ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).Devuelve / 2))
        End If
    End If
    If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
        If ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).Devuelve > 0 And UserList(UserIndex).Stats.MinHP > 0 Then
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).Devuelve)
            Call SendData(ToIndex, UserIndex, 0, "||" & UserList(tempChr).Name & " te ha devuelto " & Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).Devuelve) & " puntos de vida." & FONTTYPE_FIGHT)
            daño = daño - Porcentaje(daño, (ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).Devuelve / 2))
        End If
    End If
    ' [/GS]
    
    ' [GS] NoKO
    If NoKO = True And HayTorneo = True And UserList(tempChr).Pos.Map = MapaDeTorneo Then
        ' Si el No KO esta activado
        If UserList(tempChr).Stats.MaxHP <= daño Then daño = UserList(tempChr).Stats.MaxHP - 1
        ' Si supuestamente lo mata de una
        ' El ser no KO hace al random la vida con la que quedara la victima
    End If
    ' [/GS]
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    Miron = "InfoHechizo"
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
    Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    b = True

    Miron = "Check Die"
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr)
    Else
        Call SendUserStatsBox(tempChr)
    End If
    
    Miron = "Check Die Devolcion"
    ' [GS] Muere el Userindex? por la devolucion de la magia
    If UserList(UserIndex).Stats.MinHP < 1 Then
        'Call ContarMuerte(Userindex, tempChr) ' si lo pongo nos hacemos crimis :S
        UserList(UserIndex).Stats.MinHP = 0
        Call ActStats(UserIndex, tempChr)
        Call UserDie(UserIndex)
    Else
        Call SendUserStatsBox(UserIndex)
    End If
    ' [/GS]
    Miron = "UserAtacaUser"
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

'Mana
If Hechizos(h).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, daño, UserList(tempChr).Stats.MaxMAN)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT_YO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If

'Stamina
If Hechizos(h).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinSta, daño, _
         UserList(tempChr).Stats.MaxSta)
    If UserIndex <> tempChr Then
         Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT_YO)
    End If
    b = True
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    

    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
End If
Miron = "AutoComentarista"
' [GS] AutoComentarista
If HayTorneo = True And AutoComentarista = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    If tempChr <> UserIndex Then
        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " tira " & Hechizos(h).nombre & " sobre " & UserList(tempChr).Name & FONTTYPE_INFO)
    Else
        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " se tira " & Hechizos(h).nombre & FONTTYPE_INFO)
    End If
End If
' [/GS]

Miron = "no hay :S"

Exit Sub

Errores:
    Call LogError("Error en HechizoPropUsuario - Err " & Err.Number & " - " & Err.Description & " : " & Miron)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).nombre)

Else

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub


Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts
    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, BandidoPts, MAXREP)
    Call SendData(ToIndex, UserIndex, 0, "PN")
    If Criminal(UserIndex) Then If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
    
    
End Sub
