Attribute VB_Name = "Acciones"

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

'Sub XEstadoX(ByVal Tipo As Integer)
'On Error Resume Next
'If frmGeneral.Visible = True Then
'    Select Case Tipo
'        Case 0  ' Apagado
'            If frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Offline" Then Exit Sub
'            frmGeneral.Toolbar1.Buttons(6).Image = 37
'            frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Offline"
'        Case 1  ' Encendido
'            If frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Online" Then Exit Sub
'            frmGeneral.Toolbar1.Buttons(6).Image = 35
'            frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Online"
'        Case 2  ' User
'            frmGeneral.Toolbar1.Buttons(6).Image = 36
'            frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Online"
'        Case 3  ' Backup
'            If frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Detenido" Then Exit Sub
'            frmGeneral.Toolbar1.Buttons(6).Image = 37
'            frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: Detenido"
'        Case Else 'ni idea
'            If frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: ERROR" Then Exit Sub
'            frmGeneral.Toolbar1.Buttons(6).Image = 36
'            frmGeneral.Toolbar1.Buttons(6).Caption = "Estado: ERROR"
'    End Select
'End If
'End Sub

Sub Accion(ByVal Userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
On Error Resume Next

'¿Posicion valida?
If InMapBounds(Map, x, y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    
    

        
    '¿Es un obj?
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        UserList(Userindex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
        
        Select Case ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ObjType
            
            Case OBJTYPE_PUERTAS 'Es una puerta
                Call AccionParaPuerta(Map, x, y, Userindex)
            Case OBJTYPE_CARTELES 'Es un cartel
                Call AccionParaCartel(Map, x, y, Userindex)
            Case OBJTYPE_FOROS 'Foro
                Call AccionParaForo(Map, x, y, Userindex)
            Case OBJTYPE_LEÑA 'Leña
                If MapData(Map, x, y).OBJInfo.ObjIndex = FOGATA_APAG Then
                    Call AccionParaRamita(Map, x, y, Userindex)
                End If
            
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        UserList(Userindex).flags.TargetObj = MapData(Map, x + 1, y).OBJInfo.ObjIndex
        Call SendData(ToIndex, Userindex, 0, "SELE" & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, y, Userindex)
            
        End Select
' [OLD]
'    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
'        UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex
'        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
'        Select Case ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType
'
'            Case 6 'Es una puerta
'                Call AccionParaPuerta(Map, x + 1, y + 1, UserIndex)
'
'        End Select
'    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
'        UserList(UserIndex).flags.TargetObj = MapData(Map, x, y + 1).OBJInfo.ObjIndex
'        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
'        Select Case ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).ObjType
'
'            Case 6 'Es una puerta
'                Call AccionParaPuerta(Map, x, y + 1, UserIndex)
'
'        End Select
'
'    Else
'        UserList(UserIndex).flags.TargetNpc = 0
'        UserList(UserIndex).flags.TargetNpcTipo = 0
'        UserList(UserIndex).flags.TargetUser = 0
'        UserList(UserIndex).flags.TargetObj = 0
'        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
'    End If
'
'   End If
' [/OLD]
        
' [NEW]
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        UserList(Userindex).flags.TargetObj = MapData(Map, x + 1, y).OBJInfo.ObjIndex
        Call SendData(ToIndex, Userindex, 0, "SELE" & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, y, Userindex)
            
        End Select
    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(Userindex).flags.TargetObj = MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, Userindex, 0, "SELE" & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, y + 1, Userindex)
            
        End Select
    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(Userindex).flags.TargetObj = MapData(Map, x, y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, Userindex, 0, "SELE" & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x, y + 1, Userindex)
            
        End Select
    ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = 1 Then
             ' Sacerdote
             If UserList(Userindex).flags.Muerto <> 1 Then Exit Sub
            If Distancia(UserList(Userindex).Pos, Npclist(MapData(Map, x, y).NpcIndex).Pos) > 10 Then
                Call SendData(ToIndex, Userindex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
                Exit Sub
            End If
            If FileExist(CharPath & UCase$(UserList(Userindex).Name) & ".chr", vbNormal) = False Then
                Call SendData(ToIndex, Userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                CloseSocket (Userindex)
                Exit Sub
            End If
            If (GetTickCount - Npclist(MapData(Map, x, y).NpcIndex).Stats.LastEntrenar) > 1200 Then
                Call RevivirUsuario(Userindex)
                Npclist(MapData(Map, x, y).NpcIndex).Stats.LastEntrenar = GetTickCount
                Call SendData(ToIndex, Userindex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, Userindex, 0, "||El sacerdote esta recargando sus energias." & FONTTYPE_INFO)
            End If
    ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = 4 Then
            If Distancia(Npclist(MapData(Map, x, y).NpcIndex).Pos, UserList(Userindex).Pos) > 3 Then
                Call SendData(ToIndex, Userindex, 0, "||Estas demasiado lejos de la boveda." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call IniciarDeposito(Userindex)
    ElseIf Npclist(MapData(Map, x, y).NpcIndex).Comercia = 1 Then
            ' Es comerciante
            If Distancia(Npclist(MapData(Map, x, y).NpcIndex).Pos, UserList(Userindex).Pos) > 3 Then
                Call SendData(ToIndex, Userindex, 0, "||Estas demasiado lejos del comerciante." & FONTTYPE_INFO)
                Exit Sub
            End If
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(Userindex)
    Else
        UserList(Userindex).flags.TargetNpc = 0
        UserList(Userindex).flags.TargetNpcTipo = 0
        UserList(Userindex).flags.TargetUser = 0
        UserList(Userindex).flags.TargetObj = 0
        Call SendData(ToIndex, Userindex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If
    
End If
' [/NEW]

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Userindex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

' [GS] No hacer fogatas en una ciudad
Select Case Map
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
            Suerte = 3
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(Userindex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA
    Obj.Amount = 1
    
    Call SendData(ToIndex, Userindex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
    Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "FO")
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, x, y)
    
    
Else
    Call SendData(ToIndex, Userindex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(Userindex).flags.Hambre = 0 And UserList(Userindex).flags.Sed = 0 Then
    Call SubirSkill(Userindex, Supervivencia)
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Userindex As Integer)
On Error Resume Next

'¿Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = ""
        auxcad = ""
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call SendData(ToIndex, Userindex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(ToIndex, Userindex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Userindex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(Userindex).Pos.x, UserList(Userindex).Pos.y, x, y) > 2) Then
    If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
                          
                     MapData(Map, x, y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).IndexAbierta
                                  
                     Call MakeObj(ToMap, 0, Map, MapData(Map, x, y).OBJInfo, Map, x, y)
                     
                     'Desbloquea
                     MapData(Map, x, y).Blocked = 0
                     MapData(Map, x - 1, y).Blocked = 0
                     
                     'Bloquea todos los mapas
                     Call Bloquear(ToMap, 0, Map, Map, x, y, 0)
                     Call Bloquear(ToMap, 0, Map, Map, x - 1, y, 0)
                     
                       
                     'Sonido
                     SendData ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(ToIndex, Userindex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, x, y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).IndexCerrada
                
                Call MakeObj(ToMap, 0, Map, MapData(Map, x, y).OBJInfo, Map, x, y)
                
                
                MapData(Map, x, y).Blocked = 1
                MapData(Map, x - 1, y).Blocked = 1
                
                
                Call Bloquear(ToMap, 0, Map, Map, x - 1, y, 1)
                Call Bloquear(ToMap, 0, Map, Map, x, y, 1)
                
                SendData ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(Userindex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
    Else
        Call SendData(ToIndex, Userindex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
    End If
Else
    Call SendData(ToIndex, Userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Userindex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ObjType = 8 Then
  
  If Len(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Texto) > 0 Then
       Call SendData(ToIndex, Userindex, 0, "MCAR" & _
        ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Texto & _
        Chr(176) & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

