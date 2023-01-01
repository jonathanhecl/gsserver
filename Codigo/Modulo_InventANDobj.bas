Attribute VB_Name = "InvNpc"
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj) As WorldPos
On Error GoTo errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(Pos, NuevaPos)
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
          Call MakeObj(ToMap, 0, Pos.Map, _
          Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
          TirarItemAlPiso = NuevaPos
    End If
    
    MapInfo(Pos.Map).Objs = True

Exit Function
errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef Npc As Npc)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next
Dim MiObj As Obj
' [GS] Tira equipo
If Npc.TiraEquip = True Then
    If Npc.Equip.Arma > 0 Then
        MiObj.Amount = 1
        MiObj.ObjIndex = Npc.Equip.Arma
        Call TirarItemAlPiso(Npc.Pos, MiObj)
    End If
    If Npc.Equip.Escudo > 0 Then
        MiObj.Amount = 1
        MiObj.ObjIndex = Npc.Equip.Escudo
        Call TirarItemAlPiso(Npc.Pos, MiObj)
    End If
    If Npc.Equip.Casco > 0 Then
        MiObj.Amount = 1
        MiObj.ObjIndex = Npc.Equip.Casco
        Call TirarItemAlPiso(Npc.Pos, MiObj)
    End If
End If
' [/GS]

If Npc.Invent.NroItems > 0 Then
    
    Dim i As Byte

    
    For i = 1 To MAX_INVENTORY_SLOTS
    
        If Npc.Invent.Object(i).ObjIndex > 0 Then
              MiObj.Amount = Npc.Invent.Object(i).Amount
              MiObj.ObjIndex = Npc.Invent.Object(i).ObjIndex
              Call TirarItemAlPiso(Npc.Pos, MiObj)
        End If
      
    Next i

End If

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)


Dim i As Integer
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
On Error Resume Next
'Devuelve la cantidad original del obj de un npc
Dim Leer As clsLeerInis

Dim ln As String, npcfile As String
Dim i As Integer
 
If Npclist(NpcIndex).Numero > 499 Then
    Set Leer = LeerNPCsHostiles
Else
    Set Leer = LeerNPCs
End If

For i = 1 To MAX_INVENTORY_SLOTS
    ln = Leer.DarValor("NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next i
                   
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error Resume Next

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next i

Npclist(NpcIndex).InvReSpawn = 0

End Sub



Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).ObjType = OBJTYPE_LLAVES Then
    'If 1 = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                   
                   Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
                   Npclist(NpcIndex).Invent.Object(Slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                   If Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
                        Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
                    Else
                        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
                    End If
            End If
            
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    
    
    
    End If
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
On Error GoTo Fallo
'Vuelve a cargar el inventario del npc NpcIndex
Dim Leer As clsLeerInis
Dim LoopC As Integer
Dim ln As String

Dim npcfile As String


If Npclist(NpcIndex).Numero > 499 Then
        Set Leer = LeerNPCsHostiles
Else
        Set Leer = LeerNPCs
End If

Npclist(NpcIndex).Invent.NroItems = val(Leer.DarValor("NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

If Npclist(NpcIndex).Invent.NroItems <= 0 Then Exit Sub

For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = Leer.DarValor("NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

Exit Sub
Fallo:
    ' 0.12b3
    Call LogError("CargarInvent - Err: " & Err.Number)
End Sub


