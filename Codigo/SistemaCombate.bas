Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

Option Explicit

Public Const MAXDISTANCIAARCO = 12




Function ModAgilidad(ByVal UserIndex As Integer, ByVal Damange As Double) As Double
On Error Resume Next
ModAgilidad = Damange
' [GS] Arma que modifica magia
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Agilidad) = True Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Agilidad > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic & FONTTYPE_INFX)
            ModAgilidad = ModAgilidad * ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Agilidad
        End If
    End If
End If
' [GS] Accesorios que modifican magia
If UserList(UserIndex).Invent.Accesorio1EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Agilidad) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Agilidad > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic & FONTTYPE_INFX)
            ModAgilidad = ModAgilidad * ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Agilidad
        End If
    End If
End If
If UserList(UserIndex).Invent.Accesorio2EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Agilidad) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Agilidad > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic & FONTTYPE_INFX)
            ModAgilidad = ModAgilidad * ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Agilidad
        End If
    End If
End If
End Function

Function ModPoder(ByVal UserIndex As Integer, ByVal Damange As Double) As Double
On Error Resume Next
ModPoder = Damange
' [GS] Arma que modifica magia
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Poder) = True Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Poder > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Magic & FONTTYPE_INFX)
            ModPoder = ModPoder * ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Poder
        End If
    End If
End If
' [GS] Accesorios que modifican magia
If UserList(UserIndex).Invent.Accesorio1EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Poder) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Poder > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Magic & FONTTYPE_INFX)
            ModPoder = ModPoder * ObjData(UserList(UserIndex).Invent.Accesorio1EqpObjIndex).Poder
        End If
    End If
End If
If UserList(UserIndex).Invent.Accesorio2EqpObjIndex > 0 Then
    If IsNumeric(ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Poder) = True Then
        If ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Poder > 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||Magic:" & ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Magic & FONTTYPE_INFX)
            ModPoder = ModPoder * ObjData(UserList(UserIndex).Invent.Accesorio2EqpObjIndex).Poder
        End If
    End If
End If
End Function


Function ModificadorEvasion(ByVal clase As Byte) As Single

Select Case clase
    Case CLASS_GUERRERO
        ModificadorEvasion = 1
    Case CLASS_CAZADOR
        ModificadorEvasion = 0.9
    Case CLASS_PALADIN
        ModificadorEvasion = 0.9
    Case CLASS_BANDIDO
        ModificadorEvasion = 0.9
    Case CLASS_ASESINO
        ModificadorEvasion = 1.1
    Case CLASS_PIRATA
        ModificadorEvasion = 0.9
    Case CLASS_LADRON
        ModificadorEvasion = 1.1
    Case CLASS_BARDO
        ModificadorEvasion = 1.1
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As Byte) As Single
Select Case clase
    Case CLASS_GUERRERO
        ModificadorPoderAtaqueArmas = 1
    Case CLASS_CAZADOR
        ModificadorPoderAtaqueArmas = 0.8
    Case CLASS_PALADIN
        ModificadorPoderAtaqueArmas = 0.85
    Case CLASS_ASESINO
        ModificadorPoderAtaqueArmas = 0.85
    Case CLASS_PIRATA
        ModificadorPoderAtaqueArmas = 0.8
    Case CLASS_LADRON
        ModificadorPoderAtaqueArmas = 0.75
    Case CLASS_BANDIDO
        ModificadorPoderAtaqueArmas = 0.75
    Case CLASS_CLERIGO
        ModificadorPoderAtaqueArmas = 0.7
    Case CLASS_BARDO
        ModificadorPoderAtaqueArmas = 0.7
    Case CLASS_DRUIDA
        ModificadorPoderAtaqueArmas = 0.7
    Case CLASS_PESCADOR
        ModificadorPoderAtaqueArmas = 0.6
    Case CLASS_LEÑADOR
        ModificadorPoderAtaqueArmas = 0.6
    Case CLASS_MINERO
        ModificadorPoderAtaqueArmas = 0.6
    Case CLASS_HERRERO
        ModificadorPoderAtaqueArmas = 0.6
    Case CLASS_CARPINTERO
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.5
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As Byte) As Single
Select Case clase
    Case CLASS_GUERRERO
        ModificadorPoderAtaqueProyectiles = 0.8
    Case CLASS_CAZADOR
        ModificadorPoderAtaqueProyectiles = 1
    Case CLASS_PALADIN
        ModificadorPoderAtaqueProyectiles = 0.75
    Case CLASS_ASESINO
        ModificadorPoderAtaqueProyectiles = 0.75
    Case CLASS_PIRATA
        ModificadorPoderAtaqueProyectiles = 0.75
    Case CLASS_LADRON
        ModificadorPoderAtaqueProyectiles = 0.8
    Case CLASS_BANDIDO
        ModificadorPoderAtaqueProyectiles = 0.8
    Case CLASS_CLERIGO
        ModificadorPoderAtaqueProyectiles = 0.7
    Case CLASS_BARDO
        ModificadorPoderAtaqueProyectiles = 0.7
    Case CLASS_DRUIDA
        ModificadorPoderAtaqueProyectiles = 0.75
    Case CLASS_PESCADOR
        ModificadorPoderAtaqueProyectiles = 0.65
    Case CLASS_LEÑADOR
        ModificadorPoderAtaqueProyectiles = 0.7
    Case CLASS_MINERO
        ModificadorPoderAtaqueProyectiles = 0.65
    Case CLASS_HERRERO
        ModificadorPoderAtaqueProyectiles = 0.65
    Case CLASS_CARPINTERO
        ModificadorPoderAtaqueProyectiles = 0.7
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDañoClaseArmas(ByVal clase As Byte) As Single
Select Case clase
    Case CLASS_GUERRERO
        ModicadorDañoClaseArmas = 1.1
    Case CLASS_CAZADOR
        ModicadorDañoClaseArmas = 0.9
    Case CLASS_PALADIN
        ModicadorDañoClaseArmas = 0.9
    Case CLASS_ASESINO
        ModicadorDañoClaseArmas = 0.9
    Case CLASS_LADRON
        ModicadorDañoClaseArmas = 0.8
    Case CLASS_PIRATA
        ModicadorDañoClaseArmas = 0.8
    Case CLASS_BANDIDO
        ModicadorDañoClaseArmas = 0.8
    Case CLASS_CLERIGO
        ModicadorDañoClaseArmas = 0.8
    Case CLASS_BARDO
        ModicadorDañoClaseArmas = 0.75
    Case CLASS_DRUIDA
        ModicadorDañoClaseArmas = 0.75
    Case CLASS_PESCADOR
        ModicadorDañoClaseArmas = 0.6
    Case CLASS_LEÑADOR
        ModicadorDañoClaseArmas = 0.7
    Case CLASS_MINERO
        ModicadorDañoClaseArmas = 0.75
    Case CLASS_HERRERO
        ModicadorDañoClaseArmas = 0.75
    Case CLASS_CARPINTERO
        ModicadorDañoClaseArmas = 0.7
    Case Else
        ModicadorDañoClaseArmas = 0.5
End Select

End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As Byte) As Single
Select Case clase
    Case CLASS_GUERRERO
        ModicadorDañoClaseProyectiles = 1
    Case CLASS_CAZADOR
        ModicadorDañoClaseProyectiles = 1.1
    Case CLASS_PALADIN
        ModicadorDañoClaseProyectiles = 0.8
    Case CLASS_ASESINO
        ModicadorDañoClaseProyectiles = 0.8
    Case CLASS_LADRON
        ModicadorDañoClaseProyectiles = 0.75
    Case CLASS_PIRATA
        ModicadorDañoClaseProyectiles = 0.75
    Case CLASS_BANDIDO
        ModicadorDañoClaseProyectiles = 0.75
    Case CLASS_CLERIGO
        ModicadorDañoClaseProyectiles = 0.7
    Case CLASS_BARDO
        ModicadorDañoClaseProyectiles = 0.7
    Case CLASS_DRUIDA
        ModicadorDañoClaseProyectiles = 0.75
    Case CLASS_PESCADOR
        ModicadorDañoClaseProyectiles = 0.6
    Case CLASS_LEÑADOR
        ModicadorDañoClaseProyectiles = 0.7
    Case CLASS_MINERO
        ModicadorDañoClaseProyectiles = 0.6
    Case CLASS_HERRERO
        ModicadorDañoClaseProyectiles = 0.6
    Case CLASS_CARPINTERO
        ModicadorDañoClaseProyectiles = 0.7
    Case Else
        ModicadorDañoClaseProyectiles = 0.5
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal clase As Byte) As Single

Select Case clase
    Case CLASS_GUERRERO
        ModEvasionDeEscudoClase = 1
    Case CLASS_CAZADOR
        ModEvasionDeEscudoClase = 0.8
    Case CLASS_PALADIN
        ModEvasionDeEscudoClase = 1
    Case CLASS_ASESINO
        ModEvasionDeEscudoClase = 0.8
    Case CLASS_LADRON
        ModEvasionDeEscudoClase = 0.7
    Case CLASS_BANDIDO
        ModEvasionDeEscudoClase = 1.5
    Case CLASS_PIRATA
        ModEvasionDeEscudoClase = 0.75
    Case CLASS_CLERIGO
        ModEvasionDeEscudoClase = 0.9
    Case CLASS_BARDO
        ModEvasionDeEscudoClase = 0.75
    Case CLASS_DRUIDA
        ModEvasionDeEscudoClase = 0.75
    Case CLASS_PESCADOR
        ModEvasionDeEscudoClase = 0.7
    Case CLASS_LEÑADOR
        ModEvasionDeEscudoClase = 0.7
    Case CLASS_MINERO
        ModEvasionDeEscudoClase = 0.7
    Case CLASS_HERRERO
        ModEvasionDeEscudoClase = 0.7
    Case CLASS_CARPINTERO
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal A As Single, ByVal b As Single) As Single
If A > b Then
    Minimo = b
    Else: Minimo = A
End If
End Function

Function Maximo(ByVal A As Single, ByVal b As Single) As Single
If A > b Then
    Maximo = A
    Else: Maximo = b
End If
End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(Defensa) * _
ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
Dim PoderEvasionTemp As Long

If UserList(UserIndex).Stats.UserSkills(Tacticas) < 31 Then
    PoderEvasionTemp = (UserList(UserIndex).Stats.UserSkills(Tacticas) * _
    ModificadorEvasion(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Tacticas) < 61 Then
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorEvasion(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Tacticas) < 91 Then
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorEvasion(UserList(UserIndex).clase))
Else
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorEvasion(UserList(UserIndex).clase))
End If

PoderEvasion = (PoderEvasionTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Armas) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Armas) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
    UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
    (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
   PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
   (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
      (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueWresterling(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Wresterling) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Wresterling) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
       (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueWresterling = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

' [GS] No ataca guiardias con seguro on
If UserList(UserIndex).flags.Seguro = True And Npclist(NpcIndex).TargetNPC = NPCTYPE_GUARDIAS Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes atacar Guardias, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT_YO)
    Exit Function
End If
' [/GS]

If AntiLukers = True Then
    ' [GS] Anti robo de NPC
    ' Este NPC fue atacado por alguien?
    If Npclist(NpcIndex).flags.AttackedIndex > 0 Then
        ' Su atacante aun esta cerca Y no soy Yo
        If UserList(Npclist(NpcIndex).flags.AttackedIndex).flags.SuNPC = UserList(UserIndex).flags.TargetNPC And UserIndex <> Npclist(NpcIndex).flags.AttackedIndex Then
            ' No es el dueño
            If Criminal(Npclist(NpcIndex).flags.AttackedIndex) = False Then
                If UserList(UserIndex).flags.Seguro = True Then
                    Call SendData(ToIndex, UserIndex, 0, "||Este NPC ya tiene un dueño, para atacarlo igual debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT_YO)
                    Exit Function
                ElseIf UserList(UserIndex).flags.Seguro = False Then
                    ' Se vuelve crimi
                    Call xRobar(UserIndex)
                    Call SendData(ToIndex, Npclist(NpcIndex).flags.AttackedIndex, 0, "||" & UserList(UserIndex).Name & " te ha robado a " & Npclist(NpcIndex).Name & FONTTYPE_FIGHT)
                    Npclist(NpcIndex).flags.AttackedIndex = UserIndex
                    ' Ataka
                End If
            Else
                Call SendData(ToIndex, Npclist(NpcIndex).flags.AttackedIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado a " & Npclist(NpcIndex).Name & FONTTYPE_FIGHT)
                Npclist(NpcIndex).flags.AttackedIndex = UserIndex
                ' Ataka
            End If
        Else
            ' Ataka
        End If
    Else
        Npclist(NpcIndex).flags.AttackedIndex = UserIndex
        ' Ataka
    End If
    ' [/GS]
End If



Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueArma(UserIndex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWresterling(UserIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(UserIndex, Proyectiles)
       Else
            Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wresterling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
On Error GoTo Fallo
Dim Debuger As Integer

Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long
Debuger = 1

UserEvasion = PoderEvasion(UserIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

Debuger = 2
SkillTacticas = UserList(UserIndex).Stats.UserSkills(Tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(Defensa)

Debuger = 3
'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

' [GS] Si esta lloviendo el NPC no se zarpa
If UserList(UserIndex).Stats.MinSta < 6 And Lloviendo = True Then
    ProbExito = 10
End If
' [/GS]
Debuger = 4
NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
   If NpcImpacto = False Then
      Debuger = 5
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
        Debuger = 6
         Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_ESCUDO)
         Call SendData(ToIndex, UserIndex, 0, "7")
         Call SubirSkill(UserIndex, Defensa)
      End If
   End If
End If

Exit Function
Fallo:
Call LogError("NPCImpacto - Err: " & Err.Number & " - Parte: " & Debuger)

End Function


Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
On Error GoTo CalcFallo
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata dragones?
        If Arma.SubTipo = MATADRAGONES Then ' Usa la matadragones?
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca dragon?
                '' 0.12b4
                'DañoArma = Npclist(NpcIndex).Stats.MaxHP
                'DañoMaxArma = DañoArma
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                
            Else ' Sino es dragon daño es 10
                DañoArma = UserList(UserIndex).Stats.ELV
                DañoMaxArma = 2 * UserList(UserIndex).Stats.ELV
            End If
        Else ' daño comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If Arma.SubTipo = MATADRAGONES Then
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            DañoArma = 10 ' Si usa la espada matadragones daño es 1
            DañoMaxArma = 20
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If

Else ' [GS] Wresterling
    ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
    ' Como si fuera un arma
    DañoMaxArma = CInt(UserList(UserIndex).Stats.ELV + ModifClase)
    DañoArma = RandomNumber(1, UserList(UserIndex).Stats.ELV)
    ' [/GS]
End If
' [NEW] Hiper-AO
If Inbaneable(UserList(UserIndex).Name) Then
If NpcIndex > 0 Then
    CalcularDaño = Npclist(NpcIndex).Stats.MaxHP
Else
    CalcularDaño = 999
End If
End If
' [/NEW]
DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)
CalcularDaño = ModPoder(UserIndex, (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(Fuerza) - 15))) + DañoUsuario) * ModifClase))
If MapaAgite = UserList(UserIndex).Pos.Map Then
    CalcularDaño = CalcularDaño + Porcentaje(CalcularDaño, 2)
End If
Exit Function
CalcFallo:
    LogError ("Error en CalcularDaño - Err " & Err.Number & " - " & Err.Description)
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo Errores
Dim daño As Long
Dim Calculo As Long ' Hiper-AO

' [GS] No ataca guiardias con seguro on
If UserList(UserIndex).flags.Seguro = True And Npclist(NpcIndex).TargetNPC = NPCTYPE_GUARDIAS Then
    Call SendData(ToIndex, UserIndex, 0, "||No podes a un Guardia con el seguro activado." & FONTTYPE_INFO)
    Exit Sub
End If
' [/GS]


daño = CalcularDaño(UserIndex, NpcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(UserIndex).flags.Navegando = 1 Then _
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

daño = daño - Npclist(NpcIndex).Stats.Def

If daño < 0 Then daño = 0

' [GS] Espadas paralizantes
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Paraliza > 0 And _
        RandomNumber(1, 100) <= ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Paraliza And _
        Npclist(NpcIndex).flags.Paralizado = 0 Then
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call SendData(ToPCArea, UserIndex, Npclist(NpcIndex).Pos.Map, "CFX" & Npclist(NpcIndex).Char.CharIndex & "," & 8 & "," & 1)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado * 2
        End If
    End If
Else
    daño = daño / 2.5
    If daño < 5 Then daño = RandomNumber(1, 5)
 ' Sino wrelstring
End If
' [/GS]

' [NEW] Hiper-AO
If daño <> 0 Then
    '[Wag]
    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)
    Calculo = (daño / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)
        '[/Wag]
        
    ' [GS] Barrera?
    If Npclist(NpcIndex).flags.BarreraEspejo > 0 Then
        If CInt(RandomNumber(1, 100)) <= Npclist(NpcIndex).flags.BarreraEspejo Then
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
    
    '[Wag](Elim by Sicarul XD)

    If daño >= Npclist(NpcIndex).Stats.MinHP Then
    '    If daño >= Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
        Calculo = (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)
    '    Else
    '    Calculo = MiNPC.GiveEXP / 2 + (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP / 2)
    '    End If
    End If
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    
    Call SendData(ToIndex, UserIndex, 0, "U2" & daño)
    
    ' [GS] Party
'    If UserList(Userindex).flags.Party > 0 And UserList(UserList(Userindex).flags.Party).flags.Muerto = 0 Then
'        Call AddtoVar(UserList(Userindex).Stats.Exp, CInt(Calculo / 2), MaxExp)
'        Call SendData(ToIndex, Userindex, 0, "||Has ganado " & CInt(Calculo / 2) & " puntos de experiencia." & FONTTYPE_FIGHT)
'        Call AddtoVar(UserList(UserList(Userindex).flags.Party).Stats.Exp, Format(Calculo / 2, "#"), MaxExp)
'        Call SendData(ToIndex, UserList(Userindex).flags.Party, 0, "||Has ganado " & CInt(Calculo / 2) & " puntos de experiencia." & FONTTYPE_FIGHT)
'        Call CheckUserLevel(UserList(Userindex).flags.Party)
'    Else
    Call GanarExp(UserIndex, Calculo, False)
'    End If
    ' [/GS]
    
    
    '[/Wag]
    If Npclist(NpcIndex).Stats.MinHP > 0 Then
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(UserIndex) Then
           Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
           Call SubirSkill(UserIndex, Apuñalar)
        End If
    End If
    

    If Npclist(NpcIndex).Stats.MinHP <= 0 Then
              
              ' Si era un Dragon perdemos la espada matadragones
              If Npclist(NpcIndex).NPCtype = DRAGON Then Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
              
              ' Para que las mascotas no sigan intentando luchar y
              ' comiencen a seguir al amo
             
              Dim j As Integer
              For j = 1 To MAXMASCOTAS
                    If UserList(UserIndex).MascotasIndex(j) > 0 Then
                        If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = SIGUE_AMO
                    End If
              Next j
              Call MuereNpc(NpcIndex, UserIndex)
            Calculo = 0
    End If
End If

' [/NEW]
' [OLD]
'Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
'
'Call SendData(ToIndex, UserIndex, 0, "U2" & daño)
'
'If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
'    If PuedeApuñalar(UserIndex) Then
'       Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
'       Call SubirSkill(UserIndex, Apuñalar)
'    End If
'End If

 
'If Npclist(NpcIndex).Stats.MinHP <= 0 Then
          
          ' Si era un Dragon perdemos la espada matadragones
'          If Npclist(NpcIndex).NPCtype = DRAGON Then Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
          
          ' Para que las mascotas no sigan intentando luchar y
          ' comiencen a seguir al amo
         
'          Dim j As Integer
'          For j = 1 To MAXMASCOTAS
'                If UserList(UserIndex).MascotasIndex(j) > 0 Then
'                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = NpcIndex Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = 0
'                    Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = SIGUE_AMO
'                End If
'          Next j
'
'          Call MuereNpc(NpcIndex, UserIndex)
'End If

Exit Sub

Errores:
    Call LogError("Error en UserDañoNPC - Err " & Err.Number & " - " & Err.Description)

End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim daño As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData



daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = daño

If UserList(UserIndex).flags.Navegando = 1 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)

Select Case Lugar
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
End Select

Call SendData(ToIndex, UserIndex, 0, "N2" & Lugar & "," & daño)

If (UserList(UserIndex).flags.Privilegios = 0 And EsAdmin(UserIndex) = False) Then
    If UserList(UserIndex).flags.PocionRepelente = False Then
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
    End If
End If

'Muere el usuario
If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call SendData(ToIndex, UserIndex, 0, "6") ' Le informamos que ha muerto ;)
    
    'Si lo mato un guardia
    If Criminal(UserIndex) And Npclist(NpcIndex).NPCtype = 2 Then
        If UserList(UserIndex).Reputacion.AsesinoRep > 0 Then
             UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep - vlASESINO / 4
             If UserList(UserIndex).Reputacion.AsesinoRep < 0 Then UserList(UserIndex).Reputacion.AsesinoRep = 0
        ElseIf UserList(UserIndex).Reputacion.BandidoRep > 0 Then
             UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO / 4
             If UserList(UserIndex).Reputacion.BandidoRep < 0 Then UserList(UserIndex).Reputacion.BandidoRep = 0
        ElseIf UserList(UserIndex).Reputacion.LadronesRep > 0 Then
             UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - vlCAZADOR / 3
             If UserList(UserIndex).Reputacion.LadronesRep < 0 Then UserList(UserIndex).Reputacion.LadronesRep = 0
        End If
        If Not Criminal(UserIndex) And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = ""
        End If
    End If
    
    Call UserDie(UserIndex)

End If

End Sub
Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
        'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = NPC_ATACA_NPC
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Sub NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
' El npc puede atacar ???
On Error GoTo Fallo
Dim Debuger As Integer
Debuger = 0

' [GS] Si es un Administrador Invisible no atacar!!!
If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
' [/GS]

' [GS] Hay consulta?
If HayConsulta = True Then
    '       Es Consejero o GM ???                       Esta en el mismo mapa que el atacante???
        If (UserList(QuienConsulta).Pos.Map = Npclist(NpcIndex).Pos.Map) Then    ' NPC?
            If Distancia(Npclist(NpcIndex).Pos, UserList(QuienConsulta).Pos) < 18 Or Distancia(UserList(UserIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then Exit Sub
        End If
End If
' [/GS]

Debuger = 1

If Npclist(NpcIndex).CanAttack = 1 Then
    Call CheckPets(NpcIndex, UserIndex)
    
    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
Else
    Exit Sub
End If

Npclist(NpcIndex).CanAttack = 0
    
Debuger = 2
   
If Npclist(NpcIndex).flags.Snd1 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd1)
        
Debuger = 3
    
If NpcImpacto(NpcIndex, UserIndex) Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(UserIndex).flags.Navegando = 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Debuger = 4
    
    Call NpcDaño(NpcIndex, UserIndex)
    '¿Puede envenenar?
    
    Debuger = 4.5
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
Else
    Call SendData(ToIndex, UserIndex, 0, "N1")
End If

Debuger = 5
'-----Tal vez suba los skills------
Call SubirSkill(UserIndex, Tacticas)

Debuger = 6
Call SendUserStatsBox(val(UserIndex))
'Controla el nivel del usuario

Debuger = 7
Call CheckUserLevel(UserIndex)

Exit Sub
Fallo:

Call LogError("NPCAtacaUser - Err: " & Err.Number & " - Parte:" & Debuger)

End Sub

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

On Error GoTo Finale
Call LogTarea("NPC daña NPC: " & val(HayConsulta) & ", " & Npclist(Atacante).Name & ", " & Npclist(Victima).Name)


Dim daño As Long
Dim ANpc As Npc, DNpc As Npc
ANpc = Npclist(Atacante)
' [NEW]
Dim Calculo As Long
DNpc = Npclist(Victima)

daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño

Calculo = (daño / Npclist(Victima).Stats.MaxHP * DNpc.GiveEXP)

Call SendData(ToIndex, ANpc.MaestroUser, 0, "U4" & daño)


If Calculo > DNpc.GiveEXP Then
    Calculo = DNpc.GiveEXP
End If

' [GS] Party
'If UserList(ANpc.MaestroUser).flags.Party > 0 And UserList(UserList(ANpc.MaestroUser).flags.Party).flags.Muerto = 0 Then
'    Call AddtoVar(UserList(ANpc.MaestroUser).Stats.Exp, Format(Calculo / 2, "#"), MaxExp)
'    Call SendData(ToIndex, ANpc.MaestroUser, 0, "||Has ganado " & Format(Calculo / 2, "#") & " puntos de experiencia." & FONTTYPE_FIGHT_MASCOTA)
'    Call AddtoVar(UserList(UserList(ANpc.MaestroUser).flags.Party).Stats.Exp, Format(Calculo / 2, "#"), MaxExp)
'    Call SendData(ToIndex, UserList(ANpc.MaestroUser).flags.Party, 0, "||Has ganado " & Format(Calculo / 2, "#") & " puntos de experiencia." & FONTTYPE_FIGHT_MASCOTA)
'    Call CheckUserLevel(UserList(ANpc.MaestroUser).flags.Party)
'Else
        Call GanarExp(ANpc.MaestroUser, Calculo, True)
' [/NEW]

' [OLD]
'daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
'Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño
' [/OLD]
If Npclist(Victima).Stats.MinHP < 1 Then

        If Npclist(Atacante).flags.AttackedBy <> "" Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If

        Call FollowAmo(Atacante)

        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
End If

Finale:

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
On Error GoTo Finale
Call LogTarea("NPC ataca NPC: " & val(HayConsulta) & ", " & Npclist(Atacante).Name & ", " & Npclist(Victima).Name)

' [GS] Hay consulta?
If HayConsulta = True Then
        '       Es Consejero o GM ???                       Esta en el mismo mapa que el atacante???
        If (UserList(QuienConsulta).Pos.Map = Npclist(Atacante).Pos.Map) Then    ' NPC?
            If Distancia(Npclist(Atacante).Pos, UserList(QuienConsulta).Pos) < 18 Or Distancia(Npclist(Victima).Pos, UserList(QuienConsulta).Pos) < 18 Then Exit Sub
        End If
End If
' [/GS]

' El npc puede atacar ???
If Npclist(Atacante).CanAttack = 1 Then
       Npclist(Atacante).CanAttack = 0
       Npclist(Victima).TargetNPC = Atacante
       Npclist(Victima).Movement = NPC_ATACA_NPC
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then Call SendData(ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & Npclist(Atacante).flags.Snd1)


If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & Npclist(Victima).flags.Snd2)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & SND_IMPACTO)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & SOUND_SWING)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SOUND_SWING)
    End If
End If

Finale:

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)


Call CheckPets(NpcIndex, UserIndex)

If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > MAXDISTANCIAARCO Then
   Call SendData(ToIndex, UserIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_FIGHT_YO)
   Exit Sub
End If
' [GS]
If HayTorneo = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    ' Hay torneo y esta en el.
    ' entonces si puede atacar
Else
    If UserList(UserIndex).Faccion.ArmadaReal = 1 And Npclist(NpcIndex).MaestroUser <> 0 Then
        If Not Criminal(Npclist(NpcIndex).MaestroUser) Then
            Call SendData(ToIndex, UserIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos y sus macotas." & FONTTYPE_WARNING)
            Exit Sub
        End If
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 And Npclist(NpcIndex).MaestroUser <> 0 And LegionNoSeAtacan = True Then
        If UserList(Npclist(NpcIndex).MaestroUser).Faccion.FuerzasCaos = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Los soldados de la Legion Oscura tienen prohibido atacarse entre intregrantes y sus macotas." & FONTTYPE_WARNING)
            Exit Sub
        End If
    End If
End If

Call NpcAtacado(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    Else
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_IMPACTO2)
    End If
    
    
    
    
    Call UserDañoNpc(UserIndex, NpcIndex)
   
Else
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
    Call SendData(ToIndex, UserIndex, 0, "U1")
End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
On Error GoTo Fallo
Dim Debuger As Integer
Debuger = 1
If UserList(UserIndex).flags.PuedeAtacar = 1 Then
    
    
    'Quitamos stamina
    If UserList(UserIndex).Stats.MinSta >= 10 Then
        If UserList(UserIndex).Pos.Map <> MapaAgite Then Call QuitarSta(UserIndex, RandomNumber(1, 10))
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
        Exit Sub
    End If
    Debuger = 2
    
    UserList(UserIndex).flags.PuedeAtacar = 0
    
    Dim AttackPos As WorldPos
    AttackPos = UserList(UserIndex).Pos
    Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)
    Debuger = 3
    'Exit if not legal
    If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
        ' Antes [GS] Renvio de lado
        ' Call ChangeUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        Exit Sub
    End If
    Debuger = 4
    Dim Index As Integer
    Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
    Debuger = 5
    'Look for user
    If Index > 0 Then
        Call UsuarioAtacaUsuario(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
        Call SendUserStatsBox(UserIndex)
        Call SendUserStatsBox(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
        Exit Sub
    End If
    Debuger = 6
    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
        Debuger = 7
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
               MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Pos.Map).Pk = False Then
                Call SendData(ToIndex, UserIndex, 0, "||No podés atacar mascotas en zonas seguras" & FONTTYPE_FIGHT_YO)
                Exit Sub
        End If
        Debuger = 8
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable = 1 Or Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 Then
            Debuger = 9
            Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)
            
        Else
            Debuger = 10
            Call SendData(ToIndex, UserIndex, 0, "||No podés atacar a este NPC" & FONTTYPE_FIGHT_YO)
        End If
        Debuger = 11
        Call SendUserStatsBox(UserIndex)
        Debuger = 12
        Exit Sub
    End If
    Debuger = 13
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
    Debuger = 14
    Call SendUserStatsBox(UserIndex)
End If

Exit Sub
Fallo:
' 0.12b4
Call LogError("UsuarioAtaca - Err " & Err.Number & "(" & Err.Description & ") - Debug: " & Debuger)

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
proyectil = ObjData(Arma).proyectil = 1

'Calculamos el poder de evasion...
UserPoderEvasion = ModAgilidad(VictimaIndex, PoderEvasion(VictimaIndex))

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_ESCUDO)
              Call SendData(ToIndex, AtacanteIndex, 0, "8")
              Call SendData(ToIndex, VictimaIndex, 0, "7")
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(AtacanteIndex, Armas)
           Else
                  Call SubirSkill(AtacanteIndex, Proyectiles)
           End If
   Else
        Call SubirSkill(AtacanteIndex, Wresterling)
   End If
End If

End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
   Call SendData(ToIndex, AtacanteIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_FIGHT_YO)
   Exit Sub
End If


If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(VictimaIndex).flags.Navegando = 0 Then Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
    
Else
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SOUND_SWING)
    Call SendData(ToIndex, AtacanteIndex, 0, "U1")
    Call SendData(ToIndex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)
End If

Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
On Error GoTo Errores
Dim ParteFallo As Integer
Dim daño As Long, antdaño As Long
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer

ParteFallo = 1
' [GS] Arma solo para NPC
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).SoloNPC = 1 Then
        Call SendData(ToIndex, AtacanteIndex, 0, "||" & ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Name & " solo sirve para atacar NPC." & FONTTYPE_FIGHT_YO)
        Exit Sub
    End If
End If
' [/GS]

' [GS] AutoComentarista
If HayTorneo = True And AutoComentarista = True And UserList(AtacanteIndex).Pos.Map = MapaDeTorneo Then
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(AtacanteIndex).Name & " ataca con " & ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Name & " a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
End If
' [/GS]

ParteFallo = 2
Dim Obj As ObjData
daño = CalcularDaño(AtacanteIndex)
antdaño = daño

If UserList(AtacanteIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If
ParteFallo = 3
If UserList(VictimaIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If
ParteFallo = 4
' [GS] Arma Paralizante
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Paraliza > 0 And _
                RandomNumber(1, 100) <= ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Paraliza And _
                UserList(VictimaIndex).flags.Paralizado = 0 Then
            If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then ' Tiene Ropa
                If CInt(RandomNumber(1, 100)) <= ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex).NoParalisis And ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex).NoParalisis > 0 Then
                    ' No paraliza
                    Call SendData(ToIndex, VictimaIndex, 0, "||" & UserList(AtacanteIndex).Name & " te ha intentado paralizar." & FONTTYPE_FIGHT)
                Else
                    Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & 8 & "," & 1)
                    UserList(VictimaIndex).flags.Paralizado = 1
                    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
                    Call SendData(ToIndex, VictimaIndex, 0, "PARADOK")
                    ' [GS] AutoComentarista
                    If HayTorneo = True And AutoComentarista = True And UserList(AtacanteIndex).Pos.Map = MapaDeTorneo Then
                        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(AtacanteIndex).Name & " paraliza a " & UserList(VictimaIndex).Name & " con el golpe de " & ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Name & FONTTYPE_INFO)
                    End If
                    ' [/GS]
                End If
            Else
                Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & 8 & "," & 1)
                UserList(VictimaIndex).flags.Paralizado = 1
                UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
                Call SendData(ToIndex, VictimaIndex, 0, "PARADOK")
                ' [GS] AutoComentarista
                If HayTorneo = True And AutoComentarista = True And UserList(AtacanteIndex).Pos.Map = MapaDeTorneo Then
                    Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(AtacanteIndex).Name & " paraliza a " & UserList(VictimaIndex).Name & " con el golpe de " & ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Name & FONTTYPE_INFO)
                End If
                ' [/GS]
            End If
        End If
End If
' [/GS]

ParteFallo = 5

Lugar = RandomNumber(1, 6)

Select Case Lugar
  
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
End Select
ParteFallo = 6
Call SendData(ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & daño & "," & UserList(VictimaIndex).Name)
Call SendData(ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & daño & "," & UserList(AtacanteIndex).Name)
ParteFallo = 7
' [GS] NoKO
If NoKO = True And HayTorneo = True And UserList(VictimaIndex).Pos.Map = MapaDeTorneo Then
    ' Si el No KO esta activado
    If UserList(VictimaIndex).Stats.MaxHP <= daño Then daño = UserList(VictimaIndex).Stats.MaxHP - 1
    ' Si supuestamente lo mata de una
    ' El ser no KO hace al random la vida con la que quedara la victima
End If
' [/GS]

ParteFallo = 8
UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
        'Si usa un arma quizas suba "Combate con armas"
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SubirSkill(AtacanteIndex, Armas)
        Else
        'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, Wresterling)
        End If
        
        Call SubirSkill(AtacanteIndex, Tacticas)
        
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
        End If
End If
ParteFallo = 9

If UserList(VictimaIndex).Stats.MinHP <= 0 Then
     
     Call ContarMuerte(VictimaIndex, AtacanteIndex)
     
     ' Para que las mascotas no sigan intentando luchar y
     ' comiencen a seguir al amo
     Dim j As Integer
     For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
     Next j

     Call ActStats(VictimaIndex, AtacanteIndex)
End If
ParteFallo = 10
' [GS] Envia el statos al oponente
Call SendUserStatsBox(VictimaIndex)
' [/GS]
ParteFallo = 11
'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)



Exit Sub


Errores:
    Call LogError("Error en UserDañoUser - Parte " & ParteFallo & " - Err " & Err.Number & " - " & Err.Description)


End Sub
' [GS]
Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

If HayTorneo = True And UserList(AttackerIndex).Pos.Map = MapaDeTorneo Then
            ' Hay torneo y esta en el mapa
            ' entonces no se vuelve criminal
Else
        If Not Criminal(AttackerIndex) And Not Criminal(VictimIndex) Then
          Call VolverCriminal(AttackerIndex)
    End If
    If Not Criminal(VictimIndex) Then
          Call AddtoVar(UserList(AttackerIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    Else
          Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) > 0 And Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 Then
        ' Si ambos tienen clanes
'        MsgBox UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName)
        If UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
            Call GiveGuildPoints(1, AttackerIndex, False)
        End If
    End If
End If

End Sub
' [/GS]

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(Victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

If Origen > 0 And Destino > 0 And Origen <= UBound(UserList) And Destino <= UBound(UserList) Then
    If MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger = TRIGGER_ZONAPELEA Or _
        MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger = TRIGGER_ZONAPELEA Then
        If (MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger) Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If
Else
    TriggerZonaPelea = TRIGGER6_AUSENTE
End If

End Function

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

' [GS] Tigger

Dim T As eTrigger6
T = TriggerZonaPelea(AttackerIndex, VictimIndex)

If T = TRIGGER6_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = TRIGGER6_PROHIBE Then
    PuedeAtacar = False
    Exit Function
End If


If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = TRIGGER_ZONASEGURA Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes pelear aqui." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes pelear aqui." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

' [/GS] Tiggers

If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Esta es una zona segura, aqui no podes atacar otros usuarios." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Se asegura que la victima no es un GM
If (UserList(VictimIndex).flags.Privilegios >= 1 Or EsAdmin(VictimIndex)) Then
    SendData ToIndex, AttackerIndex, 0, "||¡¡No podes atacar a los administradores del juego!! " & FONTTYPE_WARNING
    PuedeAtacar = False
    Exit Function
End If

' [GS]
If HayTorneo = True And UserList(AttackerIndex).Pos.Map = MapaDeTorneo Then
    ' Hay torneo y esta en el
    ' entonces puede atacar
Else
    If Not Criminal(VictimIndex) And UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then
        Call SendData(ToIndex, AttackerIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    ElseIf UserList(VictimIndex).Faccion.FuerzasCaos = 1 And UserList(AttackerIndex).Faccion.FuerzasCaos = 1 And LegionNoSeAtacan = True Then
        Call SendData(ToIndex, AttackerIndex, 0, "||Los soldados de la Legion Oscura tienen prohibido atacarse entre integrantes." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
End If
' [/GS]

If UserList(VictimIndex).flags.Muerto = 1 Then
    SendData ToIndex, AttackerIndex, 0, "||No podes atacar a un espiritu." & FONTTYPE_INFO
    PuedeAtacar = False
    Exit Function
End If

' [GS]
If MapaAgite <> UserList(AttackerIndex).Pos.Map Then
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        SendData ToIndex, AttackerIndex, 0, "||No podes atacar porque estas muerto." & FONTTYPE_INFO
        PuedeAtacar = False
        Exit Function
    End If
End If
' [/GS]

If UserList(AttackerIndex).flags.Seguro Then
        If Not Criminal(VictimIndex) Then
            Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT_YO)
            Exit Function
        End If
End If

' [GS] Hay consulta?
If HayConsulta = True Then
        '       Es Consejero o GM ???                       Esta en el mismo mapa que el atacante???
        If (UserList(VictimIndex).flags.Privilegios < 1 And EsAdmin(VictimIndex) = False) And (UserList(QuienConsulta).Pos.Map = UserList(AttackerIndex).Pos.Map) Then ' User?
            If Distancia(UserList(AttackerIndex).Pos, UserList(QuienConsulta).Pos) < 18 Or Distancia(UserList(VictimIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then
                Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a nadie en medio de una consulta!!" & FONTTYPE_FIGHT_YO)
                PuedeAtacar = False
                Exit Function
            End If
        End If
End If
' [/GS]

PuedeAtacar = True

End Function


