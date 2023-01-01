Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).ObjType <> OBJTYPE_LLAVES And _
                ObjData(ObjIndex).ObjType <> OBJTYPE_BARCOS) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

' [GS] Ahora los dioses puede usar ropas de faccion sin necesidad de pertenecer a ella
If (UserList(UserIndex).flags.Privilegios = 3 Or EsAdmin(UserIndex)) Then
    'Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(ItemIndex).Name & " es para clase " & & FONTTYPE_INFX)
    ClasePuedeUsarItem = 1
    Exit Function
End If
' [/GS]

'Call LogTarea("ClasePuedeUsarItem")
' [GS]
If UserList(UserIndex).clase = ObjData(ObjIndex).ExclusivoClase Then
    ClasePuedeUsarItem = True
    Exit Function
End If
' [/GS]

Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
On Error GoTo errhandler
' [GS]
If Cantidad < 5 Then Exit Sub
'If Cantidad > 100000 Then Exit Sub
' [/GS]
Dim vecesqtiro As Integer

' [GS] No tirar oro en una ciudad!!
Select Case UserList(UserIndex).Pos.Map
    Case Nix.Map
        Call SendData(ToIndex, UserIndex, 0, "||Esta prohibido tirar oro en una ciudad, utiliza el comando /REGALAR para darle oro a otro jugador." & FONTTYPE_WARNING)
        Exit Sub
    Case Ullathorpe.Map
        Call SendData(ToIndex, UserIndex, 0, "||Esta prohibido tirar oro en una ciudad, utiliza el comando /REGALAR para darle oro a otro jugador." & FONTTYPE_WARNING)
        Exit Sub
    Case Lindos.Map
        Call SendData(ToIndex, UserIndex, 0, "||Esta prohibido tirar oro en una ciudad, utiliza el comando /REGALAR para darle oro a otro jugador." & FONTTYPE_WARNING)
        Exit Sub
    Case Banderbill.Map
        Call SendData(ToIndex, UserIndex, 0, "||Esta prohibido tirar oro en una ciudad, utiliza el comando /REGALAR para darle oro a otro jugador." & FONTTYPE_WARNING)
        Exit Sub
End Select
' [/GS]

'SI EL USER TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            If loops = 10 Then Exit Do ' Hiper-AO
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If (UserList(UserIndex).flags.Privilegios > 0 Or AaP(UserIndex)) Then
                Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
                'Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " tiro " & MiObj.Amount & " De oro." & FONTTYPE_FIGHT)
            End If
            
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                Call LogError("Error en TirarOro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

errhandler:
    Call LogError("Error TirarORO - Err " & Err.Number & " - " & Err.Description)
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

'Quita un objeto
UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    UserList(UserIndex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).NoSePasa = True Then
    Call SendData(ToIndex, UserIndex, 0, "||Este objeto no puede ser arrojado." & FONTTYPE_INFO)
    Exit Sub
  End If
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Then
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        Obj.Amount = num
        
        MapInfo(Map).Objs = True
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If (UserList(UserIndex).flags.Privilegios > 0 Or AaP(UserIndex)) Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).Name, False)
  Else
    Call SendData(ToIndex, UserIndex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If
' [GS] Actualizar inventario
Call UpdateUserInv(True, UserIndex, 0)
' [/GS]


End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
End If

MapInfo(Map).Objs = True

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Crea un Objeto
MapData(Map, X, Y).OBJInfo = Obj
Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)

'MapInfo(Map).Objs = True

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(ToIndex, UserIndex, 0, "||No puedes cargar mas objetos." & FONTTYPE_FIGHT_YO)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        ' 0.12b4 Impide que un usuario agarre un objeto bloqueado
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = True And UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False Then Exit Sub
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
        ' [OLD]
        'If ObjData(MiObj.ObjIndex).ObjType = OBJTYPE_GUITA Then
        '    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
        '    If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
        '    Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, x, y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
        '    SendUserStatsBox (UserIndex)
        '    Exit Sub
        'End If
        ' [/OLD]
        ' [NEW] Hiper-AO
        If MiObj.ObjIndex = 12 And ModoAgarre = 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            Call SendUserStatsBox(UserIndex)
            Exit Sub
        End If
        ' [/NEW]
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'call SendData(ToIndex, Userindex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If (UserList(UserIndex).flags.Privilegios = 1 Or AaP(UserIndex)) Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, True)
        End If
        
    End If
Else
    'Call SendData(ToIndex, UserIndex, 0, "||No hay nada aqui." & FONTTYPE_INFO) ' Hiper-AO no lo trae
End If
' [GS] Actualizar inventario
'Call UpdateUserInv(True, Userindex, 0)
' [/GS]
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario


Dim Obj As ObjData
If Slot = 0 Then Exit Sub
If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case Obj.ObjType
    ' [GS] Accesorios
    Case OBJTYPE_ACCESORIO
        If UserList(UserIndex).Invent.Object(Slot).ObjIndex = UserList(UserIndex).Invent.Accesorio1EqpObjIndex Then
            UserList(UserIndex).Invent.Accesorio1EqpObjIndex = 0
            UserList(UserIndex).Invent.Accesorio1EqpSlot = 0
        Else
            UserList(UserIndex).Invent.Accesorio2EqpObjIndex = 0
            UserList(UserIndex).Invent.Accesorio2EqpSlot = 0
        End If
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
    ' [/GS]
    
    Case OBJTYPE_WEAPON

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

    Case OBJTYPE_FLECHAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case OBJTYPE_HERRAMIENTAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
    
    Case OBJTYPE_ARMOUR
        
        Select Case Obj.SubTipo
        
            Case OBJTYPE_ARMADURA
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                If Obj.Cabeza > 0 Or Obj.Cabeza = -1 Then
                    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
                End If
                Call DarCuerpoDesnudo(UserIndex)
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
            Case OBJTYPE_CASCO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Case OBJTYPE_ESCUDO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End Select
        
    
End Select

Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub
Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).MUJER = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> HOMBRE
ElseIf ObjData(ObjIndex).HOMBRE = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> MUJER
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

' [GS] Ahora los dioses puede usar ropas de faccion sin necesidad de pertenecer a ella
If (UserList(UserIndex).flags.Privilegios = 3 Or EsAdmin(UserIndex)) Then
    FaccionPuedeUsarItem = 1
    Exit Function
End If
' [/GS]

' Hiper-AO borro casi todo esto!!!!!!
If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(UserIndex) Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.ArmadaReal = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(UserIndex) Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = 1
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
     Call QuitarUserInvItem(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot).Amount)
     Call SendData(ToIndex, UserIndex, 0, "||Como no te servian, las has tirado en un basurero cercano." & FONTTYPE_INFO)
     Exit Sub
End If

If Obj.MinNivel > UserList(UserIndex).Stats.ELV Then
     Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener Nivel " & Obj.MinNivel & " o superior para poder equipartelo." & FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.ObjType
    Case OBJTYPE_ACCESORIO
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                
                ' Si ya lo tengo Equipado
                If UserList(UserIndex).Invent.Accesorio1EqpSlot = Slot Or UserList(UserIndex).Invent.Accesorio2EqpSlot = Slot Then
                    ' Lo desequipo
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                ' Busco donde equiparlo
                If UserList(UserIndex).Invent.Accesorio1EqpSlot > 0 And UserList(UserIndex).Invent.Accesorio2EqpSlot > 0 Then
                    ' Los dos accesorios se estan usando
                    Call SendData(ToIndex, UserIndex, 0, "||Ya tienes demaciados accesorios equipados." & FONTTYPE_INFO)
                ElseIf UserList(UserIndex).Invent.Accesorio1EqpSlot > 0 And UserList(UserIndex).Invent.Accesorio2EqpSlot = 0 Then
                    ' El Accesorio 2 esta libre
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.Accesorio2EqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.Accesorio2EqpSlot = Slot
                Else
                    ' El Accesorio 1 esta libre y, no hay ninguno
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.Accesorio1EqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.Accesorio1EqpSlot = Slot
                End If
    
               ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
       End If
    Case OBJTYPE_WEAPON
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.WeaponEqpSlot = Slot
                
                ' [GS] Dos manos?
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).DosManos = 1 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                End If
                ' [/GS]
                
                'Sonido
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SACARARMA)
        
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
       End If
    Case OBJTYPE_HERRAMIENTAS
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                
       ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
       End If
    Case OBJTYPE_FLECHAS
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
       End If
    
    Case OBJTYPE_ARMOUR
         
         If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         
         Select Case Obj.SubTipo
         
            Case OBJTYPE_ARMADURA
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                   
                   'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex)
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                        
                    UserList(UserIndex).Char.Body = Obj.Ropaje
                    
                    ' [GS] Cabeza
                    If Obj.Cabeza > 0 Then
                        UserList(UserIndex).Char.Head = Obj.Cabeza
                    ElseIf Obj.Cabeza = -1 Then
                        UserList(UserIndex).Char.Head = 0
                    End If
                    ' [/GS]
                        
                    UserList(UserIndex).flags.Desnudo = 0
                        
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    
                    

                ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
                End If
            Case OBJTYPE_CASCO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.CascoAnim = NingunCasco
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                    End If
                    
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.CascoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
                End If
            Case OBJTYPE_ESCUDO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
       
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                    End If
            
                    ' [GS] Dos manos?
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                        End If
                    End If
                    ' [/GS]
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                    
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No perteneses a la" & IIf(ObjData(ObjIndex).Caos = 1, "s Fuerzas del Caos", " Armada Real") & "." & FONTTYPE_INFO)
                End If
        End Select
End Select

'Actualiza
Call UpdateUserInv(True, UserIndex, 0)


Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

' [GS] Ahora los dioses puede usar ropas de faccion sin necesidad de pertenecer a ella
If (UserList(UserIndex).flags.Privilegios = 3 Or EsAdmin(UserIndex)) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(ItemIndex).Name & IIf(ObjData(ItemIndex).RazaEnana = 1, " es para enanos.", " no es para enanos.") & FONTTYPE_INFX)
    CheckRazaUsaRopa = 1
    Exit Function
End If
' [/GS]

'Verifica si la raza puede usar la ropa
If UserList(UserIndex).raza = RAZA_HUMANO Or _
   UserList(UserIndex).raza = RAZA_ELFO Or _
   UserList(UserIndex).raza = RAZA_ELFO_OSCURO Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

' [GS] Esta tirando una explocion magica??
If UserList(UserIndex).flags.TiraExp = True Then
    Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
    UserList(UserIndex).flags.TiraExp = False
End If
' [/GS]
' [GS] No meditaras
If UserList(UserIndex).flags.Meditando = True Then
    Exit Sub
End If
' [/GS]

If Obj.MinNivel > UserList(UserIndex).Stats.ELV Then
     Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener Nivel " & Obj.MinNivel & " o superior para poder usarlo." & FONTTYPE_INFO)
     Exit Sub
End If

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Call QuitarUserInvItem(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot).Amount)
    Call SendData(ToIndex, UserIndex, 0, "||Como no te servian las has tirado en un basurero cercano." & FONTTYPE_INFO)
    Call UpdateUserInv(True, UserIndex, 0)
    Exit Sub
End If

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case Obj.ObjType

    Case OBJTYPE_USEONCE
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        Call AddtoVar(UserList(UserIndex).Stats.MinHam, Obj.MinHam, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).flags.Hambre = 0
        Call EnviarHambreYsed(UserIndex)
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_COMIDA
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        
        
    Case OBJTYPE_GUITA
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
    Case OBJTYPE_WEAPON
        
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
        End If

    
        If ObjData(ObjIndex).proyectil = 1 Then
            
            
            Call SendData(ToIndex, UserIndex, 0, "T01" & Proyectiles)

           
        Else
        
            If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
            
            TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
            '¿El target-objeto es leña?
            If TargObj.ObjType = OBJTYPE_LEÑA Then
                    If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                        Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap _
                             , UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                    
                    Else
             
                    End If
            End If
            
        End If
    Case OBJTYPE_POCIONES
        ' [GS] Solo torneos
        If UserList(UserIndex).Pos.Map = MapaDeTorneo And HayTorneo = True And PotsEnTorneo = False Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes tomar pociones o usar gemas en este Torneo!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        ' [/GS]
        If UserList(UserIndex).flags.PuedeAtacar = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Debes esperar unos momentos para tomar otra pocion!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        ' Hiper-AO borro una parte del siguiente codigo, yo no lo borro :P
        If UserList(UserIndex).flags.Muerto = 1 Then
            If Obj.TipoPocion <> 8 And Obj.TipoPocion <> 9 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            If Obj.TipoPocion = 8 Or Obj.TipoPocion = 9 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas vivo!! Solo podes usar este objeto cuando estas Muerto. " & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(UserIndex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
            Case 2 'Modif la fuerza
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                AddtoVar UserList(UserIndex).Stats.MinHP, RandomNumber(Obj.MinModificador, Obj.MaxModificador), UserList(UserIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5), UserList(UserIndex).Stats.MaxMAN)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call SendData(ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 6 ' Pocion desparalizante
                If UserList(UserIndex).flags.Paralizado = 1 Then
                    UserList(UserIndex).flags.Paralizado = 0
                    Call SendData(ToIndex, UserIndex, 0, "||Tu paralizacion ha sido quitada." & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "PARADOK")
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No estas paralizado." & FONTTYPE_INFO)
                End If
            Case 7 'Poti invisible
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes hacerte invisible si estas muerto." & FONTTYPE_INFO)
                    Exit Sub
                Else
                    If UserList(UserIndex).flags.Invisible = 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "||Ya estas invisible." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call SendData(ToIndex, UserIndex, 0, "||Te has hecho invisible." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.Invisible = 1
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                End If
            Case 8 ' Pocion de Resurreccion
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Te has Resucitado a ti mismo" & FONTTYPE_INFO)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 20)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este objeto estando vivo." & FONTTYPE_INFO)
                End If
            Case 9 ' Resurreccion infinita
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||Te has Resucitado a ti mismo" & FONTTYPE_INFO)
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 20)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este objeto estando vivo." & FONTTYPE_INFO)
                End If
            Case 10 ' Fuerza y Agilidad
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 11 ' Posion de repulsion
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
                UserList(UserIndex).flags.PocionRepelente = True
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
       End Select
     Case OBJTYPE_BEBIDA
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call EnviarHambreYsed(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
    
    Case OBJTYPE_LLAVES
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.ObjType = OBJTYPE_PUERTAS Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.Clave = Obj.Clave Then
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(ToIndex, UserIndex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.Clave = Obj.Clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(ToIndex, UserIndex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
    
        Case OBJTYPE_BOTELLAVACIA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            
        Case OBJTYPE_BOTELLALLENA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            
        Case OBJTYPE_HERRAMIENTAS
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlProleta, MAXREP)
            
            Select Case ObjIndex
                Case OBJTYPE_CAÑA
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "SFC")

            End Select
        
        Case OBJTYPE_PERGAMINOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Hambre = 0 And _
               UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                
            Else
               Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
            End If
       
       Case OBJTYPE_MINERALES
           If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
           End If
           Call SendData(ToIndex, UserIndex, 0, "T01" & FundirMetal)
       
       Case OBJTYPE_INSTRUMENTOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Obj.Snd1)
       
       Case OBJTYPE_BARCOS
        If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)) And _
            UserList(UserIndex).flags.Navegando = 0) _
            Or UserList(UserIndex).flags.Navegando = 1 Then
            
            ' 0.12b3
            If NivelNavegacion > UserList(UserIndex).Stats.ELV Then
                Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener Nivel " & NivelNavegacion & " o superior para poder utilizar el barco." & FONTTYPE_INFO)
            ElseIf SkillNavegacion > UserList(UserIndex).Stats.UserSkills(Navegacion) Then
                Call SendData(ToIndex, UserIndex, 0, "||Necesitas " & SkillNavegacion & " Skills en Navegacion para poder utilizar el barco." & FONTTYPE_INFO)
            Else
            
                UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.BarcoSlot = Slot
                Call DoNavega(UserIndex, Obj)
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "||¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)
        End If
        
        'Dim MiObj As Obj
        ' [GS] Al apretar U, podemos leer un cartel :P
        Case 8 ' Es tipo Cartel
            If Len(ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Texto) > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "MCAR" & _
                ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Texto & _
                Chr(176) & ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).GrhSecundario)
            End If
        ' [/GS]
End Select

'Actualiza
Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(True, UserIndex, 0)

' [GS] AutoComentarista
If HayTorneo = True And AutoComentarista = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    If Not UltimoMensajeAuto = Obj.TipoPocion Then
        Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " utiliza " & Obj.Name & FONTTYPE_INFO)
        UltimoMensajeAuto = Obj.TipoPocion
    End If
End If
' [/GS]

'Por las dudas viteh: NO lo usa Hiper-AO! al reglon este de AntiSH
UserList(UserIndex).Counters.AntiSH = UserList(UserIndex).Counters.AntiSH + 1
End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) \ ModHerreriA(UserList(UserIndex).clase) Then
        If ObjData(ArmasHerrero(i)).ObjType = OBJTYPE_WEAPON Then
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).MinHIT & "/" & ObjData(ArmasHerrero(i)).MaxHIT & ")" & "," & ArmasHerrero(i) & ","
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(ToIndex, UserIndex, 0, "LAH" & cad$)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(ToIndex, UserIndex, 0, "OBR" & cad$)

End Sub

' [GS] Sistema especial, anti bug crea items
Function EsObjConstruible(ByVal ObjIndex As Integer) As Boolean
EsObjConstruible = False
Dim i As Integer, cad$
' Carpinteria
For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ObjIndex Then
        EsObjConstruible = True
        Exit Function
    End If
Next i
' Armaduras
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ObjIndex Then
        EsObjConstruible = True
        Exit Function
    End If
Next i
' Armas
For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ObjIndex Then
        EsObjConstruible = True
        Exit Function
    End If
Next i
End Function
' [/GS]

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).clase) Then _
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).MinDef & "/" & ObjData(ArmadurasHerrero(i)).MaxDef & ")" & "," & ArmadurasHerrero(i) & ","
Next i

Call SendData(ToIndex, UserIndex, 0, "LAR" & cad$)

End Sub

                   

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

Call TirarTodosLosItems(UserIndex)

If Tirar100kAlMorir = True Then
    Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
    ' [GS] Usar Billetera?
ElseIf UserList(UserIndex).Stats.GLD < MinBilletera Then
    If UserList(UserIndex).Stats.GLD <= 0 Then Exit Sub
    Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
    ' [/GS]
End If
End Sub

Public Function ItemSeCae(ByVal Index As Integer, ByVal UserIndex As Integer) As Boolean

ItemSeCae = ObjData(Index).ObjType <> OBJTYPE_LLAVES And _
            ObjData(Index).ObjType <> OBJTYPE_BARCOS And _
            ObjData(Index).TipoPocion <> 8 And _
            ObjData(Index).TipoPocion <> 9 And _
            ObjData(Index).NoSeCae <> 1 And _
            ObjData(Index).NoSePasa <> 1
            
If ItemSeCae = False Then Exit Function

' [GS] Se caen los items
If NoSeCaenItems = True Then Exit Function
' [/GS]

If ObjData(Index).Real = 1 And UserList(UserIndex).Faccion.ArmadaReal = 0 Then
    ' El objeto es Real y es no pertenece a la armada
    ItemSeCae = False
    Exit Function
Else
    ItemSeCae = True
    Exit Function
End If

If ObjData(Index).Caos = 1 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
    ' El objeto es Caos y no pertenece a las Fuerzas de Caos
    ItemSeCae = False
    Exit Function
Else
    ItemSeCae = True
    Exit Function
End If

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

'Call LogTarea("Sub TirarTodosLosItems")

Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS

  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex, UserIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
  
Next i

End Sub


Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex, UserIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
Next i

' [GS] Si es NW crimi tira ORO
If Criminal(UserIndex) = True Then Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
' [/GS]

End Sub



