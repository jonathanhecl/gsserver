Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos


Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
' [GS]
'Public Const ExpAlUnirse = 250000
'Public Const ExpX100 = 500000
Public ExpAlUnirse
Public ExpX100
' [/GS]

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

'If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
'    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
'    Exit Sub
'End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < ParaArmada Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos " & ParaArmada & " criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 18 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 18!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 5 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

' 0.12b1
If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas reales demasiadas veces!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

' [GS]
UserList(UserIndex).Faccion.FuerzasCaos = 0
' [/GS]
UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ RecompensaXArmada

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If UserList(UserIndex).clase = CLASS_MAGO Then
           If (UserList(UserIndex).raza) = RAZA_ENANO Or _
              (UserList(UserIndex).raza) = RAZA_GNOMO Then
                  MiObj.ObjIndex = TunicaMagoImperialEnanos
           Else
                  MiObj.ObjIndex = TunicaMagoImperial
           End If
    ElseIf (UserList(UserIndex).clase) = CLASS_GUERRERO Or _
           (UserList(UserIndex).clase) = CLASS_CAZADOR Or _
           (UserList(UserIndex).clase) = CLASS_PALADIN Or _
           (UserList(UserIndex).clase) = CLASS_BANDIDO Or _
           (UserList(UserIndex).clase) = CLASS_ASESINO Then
              If (UserList(UserIndex).raza) = RAZA_ENANO Or _
                 (UserList(UserIndex).raza) = RAZA_GNOMO Then
                  MiObj.ObjIndex = ArmaduraImperial3
              Else
                  MiObj.ObjIndex = ArmaduraImperial1
              End If
    Else
              If (UserList(UserIndex).raza) = RAZA_ENANO Or _
                 (UserList(UserIndex).raza) = RAZA_GNOMO Then
                  MiObj.ObjIndex = ArmaduraImperial3
              Else
                  MiObj.ObjIndex = ArmaduraImperial2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpAlUnirse, MaxExp)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoReal(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CriminalesMatados \ RecompensaXArmada = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata " & RecompensaXArmada & " criminales mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpX100, MaxExp)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.ArmadaReal = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.FuerzasCaos = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las fuerzas del caos!!!." & FONTTYPE_FIGHT)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Aprendiz "
    Case 1
        TituloReal = "Escudero"
    Case 2
        TituloReal = "Caballero"
    Case 3
        TituloReal = "Capitan"
    Case 4
        TituloReal = "Teniente"
    Case 5
        TituloReal = "Comandante"
    Case 6
        TituloReal = "Mariscal"
    Case 7
        TituloReal = "Senescal"
    Case 8
        TituloReal = "Protector"
    Case 9
        TituloReal = "Guardian del Bien"
    Case Else
        TituloReal = "Campeón de la Luz"
End Select

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a la Legión Oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

'If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
'    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido.!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
'    Exit Sub
'End If

If UserList(UserIndex).Faccion.CiudadanosMatados < ParaCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos " & ParaCaos & " ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

' 0.12b1
If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. Vete de aquí!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Else
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas oscuras demasiadas veces!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    End If
    Exit Sub
End If
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

' [GS]
UserList(UserIndex).Faccion.ArmadaReal = 0
' [/GS]
UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100


Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a la Legión Oscura!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If (UserList(UserIndex).clase) = CLASS_MAGO Then
                If (UserList(UserIndex).raza) = RAZA_ENANO Or _
                (UserList(UserIndex).raza) = RAZA_GNOMO Then
                    MiObj.ObjIndex = TunicaMagoCaosEnanos
                Else
                    MiObj.ObjIndex = TunicaMagoCaos
                End If
    ElseIf (UserList(UserIndex).clase) = CLASS_GUERRERO Or _
           (UserList(UserIndex).clase) = CLASS_CAZADOR Or _
           (UserList(UserIndex).clase) = CLASS_PALADIN Or _
           (UserList(UserIndex).clase) = CLASS_BANDIDO Or _
           (UserList(UserIndex).clase) = CLASS_ASESINO Then
              If (UserList(UserIndex).raza) = RAZA_ENANO Or _
                 (UserList(UserIndex).raza) = RAZA_GNOMO Then
                  MiObj.ObjIndex = ArmaduraCaos3
              Else
                  MiObj.ObjIndex = ArmaduraCaos1
              End If
    Else
              If (UserList(UserIndex).raza) = RAZA_ENANO Or _
                 (UserList(UserIndex).raza) = RAZA_GNOMO Then
                  MiObj.ObjIndex = ArmaduraCaos3
              Else
                  MiObj.ObjIndex = ArmaduraCaos2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpAlUnirse, MaxExp)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CiudadanosMatados \ RecompensaXCaos = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata " & RecompensaXCaos & " ciudadanos mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpX100, MaxExp)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
End If


End Sub

Public Sub ExpulsarCaos(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.FuerzasCaos = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de la legión oscura!!!." & FONTTYPE_FIGHT)
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Esbirro"
    Case 1
        TituloCaos = "Servidor de las Sombras"
    Case 2
        TituloCaos = "Acólito"
    Case 3
        TituloCaos = "Guerrero Sombrío"
    Case 4
        TituloCaos = "Sanguinario"
    Case 5
        TituloCaos = "Caballero de la Oscuridad"
    Case 6
        TituloCaos = "Condenado"
    Case 7
        TituloCaos = "Heraldo Impío"
    Case 8
        TituloCaos = "Corruptor"
    Case Else
        TituloCaos = "Devorador de Almas"
End Select


End Function

