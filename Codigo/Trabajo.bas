Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal Userindex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim Res As Integer

If UserList(Userindex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Exit Sub
End If

If (UserList(Userindex).clase) <> CLASS_LADRON Then Suerte = Suerte + 50

Res = RandomNumber(1, Suerte)

If Res > 9 Then
   UserList(Userindex).flags.Oculto = 0
   UserList(Userindex).flags.Invisible = 0
   Call SendData(ToMap, 0, UserList(Userindex).Pos.Map, "NOVER" & UserList(Userindex).Char.CharIndex & ",0")
   Call SendData(ToIndex, Userindex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub
Public Sub DoOcultarse(ByVal Userindex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim Res As Integer

If UserList(Userindex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Suerte = 7
End If

If (UserList(Userindex).clase) <> CLASS_LADRON Then Suerte = Suerte + 50

Res = RandomNumber(1, Suerte)

If Res <= 5 Then
   UserList(Userindex).flags.Oculto = 1
   UserList(Userindex).flags.Invisible = 1
   Call SendData(ToMap, 0, UserList(Userindex).Pos.Map, "NOVER" & UserList(Userindex).Char.CharIndex & ",1")
   Call SendData(ToIndex, Userindex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
   Call SubirSkill(Userindex, Ocultarse)
Else
    If Not UserList(Userindex).flags.UltimoMensaje = 22 Then
        Call SendData(ToIndex, Userindex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 22
    End If
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal Userindex As Integer, ByRef Barco As ObjData)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(Userindex).clase)

'If UserList(UserIndex).Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
'    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
'    Call SendData(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
'    Exit Sub
'End If

If UserList(Userindex).flags.Navegando = 0 Then
    
    UserList(Userindex).Char.Head = 0
    
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Body = Barco.Ropaje
    Else
        UserList(Userindex).Char.Body = iFragataFantasmal
    End If
    
    UserList(Userindex).Char.ShieldAnim = NingunEscudo
    UserList(Userindex).Char.WeaponAnim = NingunArma
    UserList(Userindex).Char.CascoAnim = NingunCasco
    UserList(Userindex).flags.Navegando = 1
    
Else
    
    UserList(Userindex).flags.Navegando = 0
    
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
        
        If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(Userindex)
        End If
            
        If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(Userindex).Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(Userindex).Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(Userindex).Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(Userindex).Char.Body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
    End If

End If

Call ChangeUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
Call SendData(ToIndex, Userindex, 0, "NAVEG")

End Sub

Public Sub FundirMineral(ByVal Userindex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(Userindex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(Userindex).flags.TargetObjInvIndex).MinSkill <= UserList(Userindex).Stats.UserSkills(Mineria) / ModFundicion(UserList(Userindex).clase) Then
        Call DoLingotes(Userindex)
   Else
        Call SendData(ToIndex, Userindex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal Userindex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(Userindex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal Userindex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(Userindex, i)
        
        UserList(Userindex).Invent.Object(i).Amount = UserList(Userindex).Invent.Object(i).Amount - Cant
        If (UserList(Userindex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(Userindex).Invent.Object(i).Amount)
            UserList(Userindex).Invent.Object(i).Amount = 0
            UserList(Userindex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, Userindex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i
End Function

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, Userindex)
End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, Userindex) Then
                    Call SendData(ToIndex, Userindex, 0, "||No tenes suficientes madera." & FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex) Then
                    Call SendData(ToIndex, Userindex, 0, "||No tenes suficientes lingotes de hierro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex) Then
                    Call SendData(ToIndex, Userindex, 0, "||No tenes suficientes lingotes de plata." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingP, Userindex) Then
                    Call SendData(ToIndex, Userindex, 0, "||No tenes suficientes lingotes de oro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex) And UserList(Userindex).Stats.UserSkills(Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
' [GS] No construiras nada chit
If EsObjConstruible(ItemIndex) = False Then Exit Sub
' [/GS]


'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(Userindex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(Userindex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).ObjType = OBJTYPE_WEAPON Then
        Call SendData(ToIndex, Userindex, 0, "||Has construido el arma!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ESCUDO Then
        Call SendData(ToIndex, Userindex, 0, "||Has construido el escudo!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_CASCO Then
        Call SendData(ToIndex, Userindex, 0, "||Has construido el casco!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ARMOUR Then
        Call SendData(ToIndex, Userindex, 0, "||Has construido la armadura!." & FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    Call SubirSkill(Userindex, Herreria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
' [GS] No construiras nada chit
If EsObjConstruible(ItemIndex) = False Then Exit Sub
' [/GS]

If CarpinteroTieneMateriales(Userindex, ItemIndex) And _
   UserList(Userindex).Stats.UserSkills(Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(Userindex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(Userindex, ItemIndex)
    Call SendData(ToIndex, Userindex, 0, "||Has construido el objeto!" & FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call SubirSkill(Userindex, Carpinteria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

End Sub

Public Sub DoLingotes(ByVal Userindex As Integer)
'    Call LogTarea("Sub DoLingotes")
    If UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount < 5 Then
              Call SendData(ToIndex, Userindex, 0, "||No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
              Exit Sub
    End If
    
    If RandomNumber(1, ObjData(UserList(Userindex).flags.TargetObjInvIndex).MinSkill) < 10 Then
                UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount = UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount - 5
                If UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount < 1 Then
                    UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount = 0
                    UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).ObjIndex = 0
                End If
                Call SendData(ToIndex, Userindex, 0, "||Has obtenido un lingote!!!" & FONTTYPE_INFO)
                Dim nPos As WorldPos
                Dim MiObj As Obj
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObjInvIndex).LingoteIndex
                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
                End If
                Call UpdateUserInv(False, Userindex, UserList(Userindex).flags.TargetObjInvSlot)
                Call SendData(ToIndex, Userindex, 0, "||¡Has obtenido un lingote!" & FONTTYPE_INFO)
    Else
        
        UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount = UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount - 5
        If UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount < 1 Then
                UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).Amount = 0
                UserList(Userindex).Invent.Object(UserList(Userindex).flags.TargetObjInvSlot).ObjIndex = 0
        End If
        Call UpdateUserInv(False, Userindex, UserList(Userindex).flags.TargetObjInvSlot)
        Call SendData(ToIndex, Userindex, 0, "||Los minerales no eran de buena calidad, no has logrado hacer un lingote." & FONTTYPE_INFO)
    End If
    
End Sub

Function ModNavegacion(ByVal clase As Byte) As Integer

Select Case clase
    Case CLASS_PIRATA
        ModNavegacion = 1
    Case CLASS_PESCADOR
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal clase As Byte) As Integer

Select Case clase
    Case CLASS_MINERO
        ModFundicion = 1
    Case CLASS_HERRERO
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As Byte) As Integer

Select Case clase
    Case CLASS_CARPINTERO
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As Byte) As Integer

Select Case clase
    Case CLASS_HERRERO
        ModHerreriA = 1
    Case CLASS_MINERO
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal clase As Byte) As Integer
Select Case clase
    Case CLASS_DRUIDA
        ModDomar = 6
    Case CLASS_CAZADOR
        ModDomar = 6
    Case CLASS_CLERIGO
        ModDomar = 7
    Case Else
        ModDomar = 10
End Select
End Function

Function CalcularPoderDomador(ByVal Userindex As Integer) As Long
CalcularPoderDomador = _
UserList(Userindex).Stats.UserAtributos(Carisma) * _
(UserList(Userindex).Stats.UserSkills(Domar) / ModDomar(UserList(Userindex).clase)) _
+ RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Carisma) / 3)
End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer
'Call LogTarea("Sub FreeMascotaIndex")
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(Userindex).MascotasIndex(j) = 0 Then
        FreeMascotaIndex = j
        Exit Function
    End If
Next j
End Function
Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(Userindex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = Userindex Then
        Call SendData(ToIndex, Userindex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, Userindex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(Userindex) Then
        Dim Index As Integer
        UserList(Userindex).NroMacotas = UserList(Userindex).NroMacotas + 1
        Index = FreeMascotaIndex(Userindex)
        UserList(Userindex).MascotasIndex(Index) = NpcIndex
        UserList(Userindex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = Userindex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(ToIndex, Userindex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
        Call SubirSkill(Userindex, Domar)
        
    Else
    
        Call SendData(ToIndex, Userindex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
        
    End If
Else
    Call SendData(ToIndex, Userindex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal Userindex As Integer)
    
    If UserList(Userindex).flags.AdminInvisible = 0 Then
        
        UserList(Userindex).flags.AdminInvisible = 1
        UserList(Userindex).flags.Invisible = 1
        UserList(Userindex).flags.OldBody = UserList(Userindex).Char.Body
        UserList(Userindex).flags.OldHead = UserList(Userindex).Char.Head
        UserList(Userindex).Char.Body = 0
        UserList(Userindex).Char.Head = 0
        If HayConsulta = True And QuienConsulta = Userindex Then
            HayConsulta = False
            Call SendData(ToAdmins, 0, 0, "||Modo Consulta DESACTIVADA" & FONTTYPE_FIGHT)
        End If
    Else
        
        UserList(Userindex).flags.AdminInvisible = 0
        UserList(Userindex).flags.Invisible = 0
        UserList(Userindex).Char.Body = UserList(Userindex).flags.OldBody
        UserList(Userindex).Char.Head = UserList(Userindex).flags.OldHead
        If HayConsulta = True = QuienConsulta = Userindex Then
            HayConsulta = False
            Call SendData(ToAdmins, 0, 0, "||Modo Consulta DESACTIVADA" & FONTTYPE_FIGHT)
        End If
    End If
    
    
    Call ChangeUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
    
End Sub
Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Userindex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj

If Not LegalPos(Map, x, y) Then Exit Sub

If MapData(Map, x, y).OBJInfo.Amount < 3 Then
    Call SendData(ToIndex, Userindex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
    Exit Sub
End If
' [GS] No hacer fogatas en una ciudad
Select Case UserList(Userindex).Pos.Map
    Case Nix.Map
        Call SendData(ToIndex, Userindex, 0, "||Esta prohibido hacer fogatas en una ciudad." & FONTTYPE_WARNING)
        Exit Sub
    Case Ullathorpe.Map
        Call SendData(ToIndex, Userindex, 0, "||Esta prohibido hacer fogatas en una ciudad." & FONTTYPE_WARNING)
        Exit Sub
    Case Lindos.Map
        Call SendData(ToIndex, Userindex, 0, "||Esta prohibido hacer fogatas en una ciudad." & FONTTYPE_WARNING)
        Exit Sub
    Case Banderbill.Map
        Call SendData(ToIndex, Userindex, 0, "||Esta prohibido hacer fogatas en una ciudad." & FONTTYPE_WARNING)
        Exit Sub
    Case Prision.Map
        Call SendData(ToIndex, Userindex, 0, "||Esta prohibido hacer fogatas en la Prision." & FONTTYPE_WARNING)
        Exit Sub
End Select
' [/GS]
If UserList(Userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(Userindex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 4
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 3
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 10 Then
            Suerte = 2
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, x, y).OBJInfo.Amount / 3
    
    If Obj.Amount > 1 Then
        Call SendData(ToIndex, Userindex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, Userindex, 0, "||Has hecho una fogata." & FONTTYPE_INFO)
    End If
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, x, y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.x = x
    Fogatita.y = y
    Call TrashCollector.Add(Fogatita)
    
Else
    Call SendData(ToIndex, Userindex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
End If

Call SubirSkill(Userindex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal Userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim Res As Integer


If UserList(Userindex).clase = CLASS_PESCADOR Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
End If

If UserList(Userindex).Stats.UserSkills(Pesca) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(Userindex).Stats.UserSkills(Pesca) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Pesca) >= 91 Then
                    Suerte = 7
End If
Res = RandomNumber(1, Suerte)

If Res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = RandomNumber(1, (UserList(Userindex).Stats.UserSkills(Pesca) * 2)) '1 ' Hiper-AO
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, Userindex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, Userindex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
End If

Call SubirSkill(Userindex, Pesca)


Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If MapInfo(UserList(VictimaIndex).Pos.Map).Pk = 1 Then Exit Sub

' [GS] Hay torneo, no robaras
If UserList(LadrOnIndex).Pos.Map = MapaDeTorneo And HayTorneo = True Then
    Call SendData(ToIndex, LadrOnIndex, 0, "||No puedes robar en un torneo!" & FONTTYPE_INFO)
    Exit Sub
End If
' [/GS]

' [GS] Hay consulta?
If HayConsulta = True Then
        If (UserList(LadrOnIndex).flags.Privilegios < 1 And EsAdmin(LadrOnIndex) = False) And (UserList(QuienConsulta).Pos.Map = UserList(LadrOnIndex).Pos.Map) Then  ' NPC?
            If Distancia(UserList(LadrOnIndex).Pos, UserList(QuienConsulta).Pos) < 18 Or Distancia(UserList(VictimaIndex).Pos, UserList(QuienConsulta).Pos) < 18 Then
                Call SendData(ToIndex, LadrOnIndex, 0, "||No puedes robar en medio de una consulta!" & FONTTYPE_FIGHT_YO)
                Exit Sub
            End If
        End If
End If
' [/GS]

If UserList(VictimaIndex).flags.Privilegios < 2 And EsAdmin(VictimaIndex) = False Then
    Dim Suerte As Integer
    Dim Res As Integer
    
       
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 91 Then
                        Suerte = 5
    End If
    Res = RandomNumber(1, Suerte)
    
    If Res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).clase) = CLASS_LADRON Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                If UserList(VictimaIndex).Stats.GLD > 1000 Then
                    N = CInt(RandomNumber(1000, UserList(VictimaIndex).Stats.GLD))
                Else
                    N = UserList(VictimaIndex).Stats.GLD
                End If
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MaxOro)
                
                Call SendData(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(ToIndex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)
    End If

    Call xRobar(LadrOnIndex)
End If


End Sub


Public Function xRobar(ByVal LadrOnIndex As Integer)
If Not Criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
End If
If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
Call AddtoVar(UserList(LadrOnIndex).Reputacion.LadronesRep, vlLadron, MAXREP)
Call SubirSkill(LadrOnIndex, Robar)

' [GS] Corrige error de mapa
Call ResetUserChar(ToMap, 0, UserList(LadrOnIndex).Pos.Map, LadrOnIndex)
' [/GS]

End Function

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).ObjType <> OBJTYPE_LLAVES And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).ObjType <> OBJTYPE_BARCOS And _
ObjData(OI).NoSeCae = False And _
ObjData(OI).NoSeVende = False And _
ObjData(OI).NoSePasa = False

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

Dim Suerte As Integer
Dim Res As Integer

If UserList(Userindex).Stats.UserSkills(Apuñalar) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(Apuñalar) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Apuñalar) >= 91 Then
                    Suerte = 5
End If
Res = RandomNumber(1, Suerte)

' [GS] Si es acesino tiene 2 chances
If UserList(Userindex).clase = CLASS_ASESINO Then
    If Res <> 3 Then
        Res = RandomNumber(1, Suerte)
    End If
End If
' [/GS]

If Res = 3 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - (daño * 1.5)
        Call SendData(ToIndex, Userindex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & CLng(daño * 1.5) & FONTTYPE_FIGHT_YO)
        Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(Userindex).Name & " por " & CLng(daño * 1.5) & FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - CLng(daño * 2)
        Call SendData(ToIndex, Userindex, 0, "||Has apuñalado la criatura por " & CLng(daño * 2) & FONTTYPE_FIGHT_YO)
        Call SubirSkill(Userindex, Apuñalar)
    End If
    
Else
    Call SendData(ToIndex, Userindex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT_YO)
End If

End Sub

Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)
UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Cantidad
If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal Userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim Res As Integer


If UserList(Userindex).clase = CLASS_LEÑADOR Then
    Call QuitarSta(Userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(Userindex, EsfuerzoTalarGeneral)
End If

If UserList(Userindex).Stats.UserSkills(Talar) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(Userindex).Stats.UserSkills(Talar) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Talar) >= 91 Then
                    Suerte = 7
End If
Res = RandomNumber(1, Suerte)

If Res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
        ' [OLD] 'MiObj.Amount = RandomNumber(1000, 5000) ' No Hiper-AO
    ' [NEW]
    If UserList(Userindex).clase = CLASS_LEÑADOR Then
        MiObj.Amount = RandomNumber(1, (UserList(Userindex).Stats.UserSkills(Talar) * 5))
    Else
        MiObj.Amount = RandomNumber(1, (UserList(Userindex).Stats.UserSkills(Talar) * 2))
    End If
    ' [/NEW]
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        
    End If
    
    Call SendData(ToIndex, Userindex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, Userindex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
End If

Call SubirSkill(Userindex, Talar)

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal Userindex As Integer)
'If UserList(UserIndex).flags.Privilegios < 2 Then
' [GS] Hay torneo, no cuenta
If UserList(Userindex).Pos.Map = MapaDeTorneo And HayTorneo = True Then Exit Sub
' [/GS]
    UserList(Userindex).Reputacion.BurguesRep = 0
    UserList(Userindex).Reputacion.NobleRep = 0
    UserList(Userindex).Reputacion.PlebeRep = 0
    Call AddtoVar(UserList(Userindex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    If UserList(Userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)
'End If
' [GS] Corrige error de mapa
Call ResetUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex)
' [/GS]

End Sub

Sub VolverCiudadano(ByVal Userindex As Integer)
' [GS] Hay torneo, no cuenta
If UserList(Userindex).Pos.Map = MapaDeTorneo And HayTorneo = True Then Exit Sub
' [/GS]
UserList(Userindex).Reputacion.LadronesRep = 0
UserList(Userindex).Reputacion.BandidoRep = 0
UserList(Userindex).Reputacion.AsesinoRep = 0
Call AddtoVar(UserList(Userindex).Reputacion.PlebeRep, vlASALTO, MAXREP)

' [GS] Corrige error de mapa
Call ResetUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex)
' [/GS]

End Sub


Public Sub DoPlayInstrumento(ByVal Userindex As Integer)

End Sub

Public Sub DoMineria(ByVal Userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim Res As Integer
Dim metal As Integer

If UserList(Userindex).clase = CLASS_MINERO Then
    Call QuitarSta(Userindex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(Userindex, EsfuerzoExcavarGeneral)
End If

If UserList(Userindex).Stats.UserSkills(Mineria) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(Mineria) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Mineria) >= 91 Then
                    Suerte = 7
End If
Res = RandomNumber(1, Suerte)

If Res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObj).MineralIndex
    
    ' [NEW]
    If UserList(Userindex).clase = CLASS_MINERO Then
        MiObj.Amount = RandomNumber(1, (UserList(Userindex).Stats.UserSkills(Mineria) * 6))
    Else
        MiObj.Amount = RandomNumber(1, (UserList(Userindex).Stats.UserSkills(Mineria) * 2))
    End If
    ' [/NEW]
    ' [OLD]
    'If UserList(UserIndex).Clase = "Minero" Then
    '    MiObj.Amount = RandomNumber(1, 6)
    'Else
    '    MiObj.Amount = 1
    'End If
    ' [/OLD]
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then _
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    
    Call SendData(ToIndex, Userindex, 0, "||¡Has extraido algunos minerales!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, Userindex, 0, "||¡No has conseguido nada!" & FONTTYPE_INFO)
End If

Call SubirSkill(Userindex, Mineria)


Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal Userindex As Integer)

UserList(Userindex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim Res As Integer
Dim Cant As Integer

If UserList(Userindex).Stats.MinMAN >= UserList(Userindex).Stats.MaxMAN Then
    Call SendData(ToIndex, Userindex, 0, "||Has terminado de meditar." & FONTTYPE_INFO)
    Call SendData(ToIndex, Userindex, 0, "MEDOK")
    UserList(Userindex).flags.Meditando = False
    UserList(Userindex).Char.FX = 0
    UserList(Userindex).Char.loops = 0
    Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(Userindex).Stats.UserSkills(Meditar) <= 10 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 20 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 30 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 40 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 50 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 60 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 70 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 80 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 90 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(Meditar) <= 100 _
   And UserList(Userindex).Stats.UserSkills(Meditar) >= 91 Then
                    Suerte = 5
End If
Res = RandomNumber(1, Suerte)

If Res = 1 Then
    Cant = Porcentaje(UserList(Userindex).Stats.MaxMAN, 3)
    Call AddtoVar(UserList(Userindex).Stats.MinMAN, Cant, UserList(Userindex).Stats.MaxMAN)
    Call SendData(ToIndex, Userindex, 0, "||¡Has recuperado " & Cant & " puntos de mana!" & FONTTYPE_INFO)
    Call SendUserStatsBox(Userindex)
    Call SubirSkill(Userindex, Meditar)
End If

End Sub




