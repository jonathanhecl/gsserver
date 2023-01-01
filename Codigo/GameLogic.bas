Attribute VB_Name = "Extra"
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

Public Sub Alerta(ByVal Texto As String)
On Error Resume Next
Dim NoTa As Boolean
If Len(Texto) < 1 Then Exit Sub
NoTa = False
If frmGeneral.Visible = False Then NoTa = True
If frmG_Alertas.Visible = False Then
    frmG_Alertas.Al.AddItem Now & " - " & Texto
    frmG_Alertas.Visible = False
Else
    frmG_Alertas.Al.AddItem Now & " - " & Texto
End If
If NoTa = True Then frmGeneral.Visible = False
End Sub


Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT
    End If
    
    If MapData(Map, X, Y).TileExit.Map > 0 Then
        '¿Es mapa de newbies?
        If MapaValido(MapData(Map, X, Y).TileExit.Map) = False Then
            If EsAdmin(UserIndex) = True Or UserList(UserIndex).flags.Privilegios > 1 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 12 & MapData(Map, X, Y).TileExit.Map Then
                    Call SendData(ToIndex, UserIndex, 0, "||Teletransporte invalido, mapa " & MapData(Map, X, Y).TileExit.Map & " inexistente." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 12 & MapData(Map, X, Y).TileExit.Map
                End If
            End If
            Exit Sub
        End If
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(ToIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                
                Call ClosestLegalPos(UserList(UserIndex).Pos, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies
            ' [GS] Aventura
            If UserList(UserIndex).flags.AV_Esta = True And UserList(UserIndex).flags.Muerto = 0 Then
                ' Esta en una aventura
                If MapData(Map, X, Y).TileExit.Map = Ullathorpe.Map Then
                    ' Si es ulla lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Tú aventura ha terminado." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Banderbill.Map Then
                    ' Si es bander lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Tú aventura ha terminado." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Nix.Map Then
                    ' Si es nix lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Tú aventura ha terminado." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Lindos.Map Then
                    ' Si es lindos lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Tú aventura ha terminado." & "~255~255~0~1~0")
                Else
                    Exit Sub ' No le deja salir caminando
                End If
            ElseIf UserList(UserIndex).Pos.Map <> MapaAventura Then
                ' No esta en el mapa de aventura
                If MapData(Map, X, Y).TileExit.Map = MapaAventura Then
                    ' Y quiere pasar al mapa de aventura
                    Exit Sub ' No le deja entrar al mapa desde afuera
                End If
            End If
            ' [/GS]
            ' [GS] Counter
            If UserList(UserIndex).flags.CS_Esta = True Then
                If MapData(Map, X, Y).TileExit.Map = Ullathorpe.Map Then
                    ' Si es ulla lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Has salido del juego." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Banderbill.Map Then
                    ' Si es bander lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Has salido del juego." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Nix.Map Then
                    ' Si es nix lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Has salido del juego." & "~255~255~0~1~0")
                ElseIf MapData(Map, X, Y).TileExit.Map = Lindos.Map Then
                    ' Si es lindos lo deja
                    Call SendData(ToIndex, UserIndex, 0, "||Has salido del juego." & "~255~255~0~1~0")
                Else
                    Exit Sub ' No le deja salir caminando
                End If
            ElseIf UserList(UserIndex).Pos.Map <> MapaCounter Then
                ' No esta en el mapa de counter
                If MapData(Map, X, Y).TileExit.Map = MapaCounter Then
                    ' Y quiere pasar al mapa counter
                    Exit Sub ' No le deja entrar al mapa desde afuera
                End If
            End If
            ' [/GS]
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents - Err " & Err.Number & " - " & Err.Description & " - User: " & UserList(UserIndex).Name)

End Sub

Function InRangoVision(ByVal UserIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
    If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

If MapInfo(Pos.Map).Cargado = False Then Exit Sub

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer

Dim UserIndex As Integer
' Espacio en el nick?
Name = Replace(Name, ".", " ")
Name = Replace(Name, "+", " ")
' [GS]
If Left(Name, 1) = "^" Or Left(Name, 1) = "[" Then Name = "GS"
' [/GS]
'¿Nombre valido?
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UCase$(Left$(UserList(UserIndex).Name, Len(Name))) = UCase$(Name)
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        UserIndex = 0
        Exit Do
    End If
    
Loop
NameIndex = UserIndex
End Function


Function IP_Index(ByVal inIP As String) As Integer
On Error GoTo local_errHand

Dim UserIndex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Do
    End If
    
Loop

IP_Index = UserIndex

Exit Function

local_errHand:
    IP_Index = 0
    

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then 'El Hiper-AO tiene un bug aqui!!!
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = EAST Then
    nX = X + 1
    nY = Y
End If

If Head = WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nX
Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else

    If MapInfo(Map).Cargado = True Then
        If Not PuedeAgua Then
              LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                         (MapData(Map, X, Y).UserIndex = 0) And _
                         (MapData(Map, X, Y).NpcIndex = 0) And _
                         (Not HayAgua(Map, X, Y))
        Else
              LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                         (MapData(Map, X, Y).UserIndex = 0) And _
                         (MapData(Map, X, Y).NpcIndex = 0) And _
                         (HayAgua(Map, X, Y))
        End If
    End If
End If

End Function



Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else
    If MapInfo(Map).Cargado = True Then
        If AguaValida = 0 Then
          LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
            (MapData(Map, X, Y).UserIndex = 0) And _
            (MapData(Map, X, Y).NpcIndex = 0) And _
            (MapData(Map, X, Y).trigger <> POSINVALIDA) _
            And Not HayAgua(Map, X, Y)
        Else
          LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
            (MapData(Map, X, Y).UserIndex = 0) And _
            (MapData(Map, X, Y).NpcIndex = 0) And _
            (MapData(Map, X, Y).trigger <> POSINVALIDA)
        End If
    End If
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(ToIndex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC
End Sub
Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

If Npclist(NpcIndex).NroExpresiones > 0 Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
End If
                    
End Sub
Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo FalloLock:

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim NpcIndex As Integer ' Hiper-AO

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y
    ' v12a12
    If EsAdmin(UserIndex) Or UserList(UserIndex).flags.Privilegios > 0 Then
        ' Si es Admin
        If MapData(Map, X, Y).TileExit.Map > 0 Then
            ' Nos muestra si hay teleport y a donde va
            If MapData(Map, X, Y).OBJInfo.ObjIndex = 378 Then
                Call SendData(ToIndex, UserIndex, 0, "||Teletransporte a Mapa: " & MapData(Map, X, Y).TileExit.Map & ", X: " & MapData(Map, X, Y).TileExit.X & ", Y: " & MapData(Map, X, Y).TileExit.Y & FONTTYPE_INFX)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Teletransporte Invisible a Mapa: " & MapData(Map, X, Y).TileExit.Map & ", X: " & MapData(Map, X, Y).TileExit.X & ", Y: " & MapData(Map, X, Y).TileExit.Y & FONTTYPE_INFX)
            End If
        End If
    End If
    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        'Call SendData(ToIndex, Userindex, 0, "||" & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Name & " - Cant: " & MapData(Map, x, y).OBJInfo.Amount & FONTTYPE_INFO)
        'UserList(Userindex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Call SendData(ToIndex, Userindex, 0, "||" & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            'UserList(Userindex).flags.TargetObj = MapData(Map, x + 1, y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            'Call SendData(ToIndex, Userindex, 0, "||" & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            'UserList(Userindex).flags.TargetObj = MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            'Call SendData(ToIndex, Userindex, 0, "||" & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            'UserList(Userindex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
        
        If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(UserList(UserIndex).flags.TargetObj).Name & " - Cant: " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.Amount & " - OBJIndex: " & UserList(UserIndex).flags.TargetObj & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(UserList(UserIndex).flags.TargetObj).Name & " - Cant: " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.Amount & FONTTYPE_INFO)
        End If
        
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
    
        ' v0.12b1
        If UserList(TempCharIndex).NoExiste = True Then Exit Sub
    
        If HayQuest = True And UserList(TempCharIndex).flags.AdminInvisible = 1 Then
            Call DoAdminInvisible(TempCharIndex)
            Call SendData(ToAll, 0, 0, "||<Quest> " & UserList(TempCharIndex).Name & " ha sido descubierto por " & UserList(UserIndex).Name & FONTTYPE_TALK & ENDC)
            Call SendData(ToIndex, UserIndex, 0, "||<Quest> Felicidades, descubriste a " & UserList(TempCharIndex).Name & FONTTYPE_TALK & ENDC)
            DoEvents
        End If
        
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Then
                        
            ' Config Click :D
            ' 0 - NICK - HP - LVL - (FACCION - CASADO) - CLAN - KILLS
            ' 1 - NICK - LVL - (FACCION - CASADO) - KILLS
            ' 2 - NICK - LVL - (FACCION - CASADO)
            ' 3 - NICK - (FACCION - CASADO)
            ' 4 - NICK - (FACCION - CASADO) - CLAN
            ' 5 - NICK - (FACCION - CASADO) - KILL
            ' 6 - NICK - CLAN
            ' 7 - NICK - LVL - CLAN
            ' 8 - NICK - LVL
            ' 9 - NICK - KILL
            ' 10 - NICK
            
            If UCase$(UserList(TempCharIndex).Name) <> "GS" Then
            
            If EsNewbie(TempCharIndex) Then
                Stat = " <NEWBIE>"
            End If
            
            '1 2 3 4 5, ok
            If ConfigClick <= 5 Then
                ' FACCION
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Armada Real> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                End If
                ' CASAMIENTO
                ' [NEW]
                If UserList(TempCharIndex).flags.Casado <> "" Then ' Hiper-AO
                    Dim ReRaRo$
                    ReRaRo$ = IIf(UserList(TempCharIndex).genero = HOMBRE, "Casado", "Casada")
                    Stat = Stat & " <" & ReRaRo$ & " con " & UserList(TempCharIndex).flags.Casado & ">"
                End If
                ' [/NEW]
            End If
            
            '1,4,6,7
            If ConfigClick = 0 Or ConfigClick = 4 Or ConfigClick = 6 Or ConfigClick = 7 Then
                If UserList(TempCharIndex).GuildInfo.GuildName <> "" Then
                    Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
                End If
            End If
            
            ' HP
            ' 0
            If ConfigClick = 0 Then
                If UserList(TempCharIndex).flags.Muerto Then
                    Stat = Stat & " [MUERTO]"
                Else
                    Stat = Stat & " [" & UserList(TempCharIndex).Stats.MinHP & "/" & UserList(TempCharIndex).Stats.MaxHP & "]"
                End If
            End If
            ' LVL
            ' 0 1 2 7 8
            If ConfigClick = 0 Or ConfigClick = 1 Or ConfigClick = 2 Or ConfigClick = 7 Or ConfigClick = 8 Then
                Stat = Stat & " [Nivel: " & UserList(TempCharIndex).Stats.ELV & "]"
            End If
            
            
            ' KILL
            ' 0 1 5 9
            If ConfigClick = 0 Or ConfigClick = 1 Or ConfigClick = 5 Or ConfigClick = 9 Then
                ' [NEW] Hiper-AO (pero no me gusta)
                If UserList(TempCharIndex).Faccion.CiudadanosMatados > 0 And UserList(TempCharIndex).Faccion.CriminalesMatados > 0 Then
                    Stat = Stat & " [ /Ciudas Matados:" & str(UserList(TempCharIndex).Faccion.CiudadanosMatados) & " /Crimis Matados:" & str(UserList(TempCharIndex).Faccion.CriminalesMatados) & " ]"
                ElseIf UserList(TempCharIndex).Faccion.CiudadanosMatados > 0 Then
                    Stat = Stat & " [ /Ciudas Matados:" & str(UserList(TempCharIndex).Faccion.CiudadanosMatados) & " ]"
                ElseIf UserList(TempCharIndex).Faccion.CriminalesMatados > 0 Then
                    Stat = Stat & " [ /Crimis Matados:" & str(UserList(TempCharIndex).Faccion.CriminalesMatados) & " ]"
                End If
                ' [/NEW]
            End If
                          
            End If


            If Len(UserList(TempCharIndex).Desc) > 1 Then
                If UCase(UserList(TempCharIndex).Name) = "GS" Then
                    Stat = "||Ves a ^[GS]^" & Stat & " - " & UserList(TempCharIndex).Desc
                Else
                    Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
                End If
            Else
                If UCase(UserList(TempCharIndex).Name) = "GS" Then
                    Stat = "||Ves a ^[GS]^" & Stat
                Else
                    'Call SendData(ToIndex, UserIndex, 0, "||Ves a " & UserList(TempCharIndex).Name & Stat)
                    Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat
                End If
            End If

            
            If AaP(TempCharIndex) Then
                If Criminal(TempCharIndex) Then
                    Stat = Stat & " <CRIMINAL>"
                Else
                    Stat = Stat & " <CIUDADANO>"
                End If
                Stat = Stat & " <APRENDIZ DE ADMINISTRADOR> ~0~185~0~1~0"
            ElseIf EsAdmin(TempCharIndex) Then
                If Criminal(TempCharIndex) Then
                    Stat = Stat & " <CRIMINAL>"
                Else
                    Stat = Stat & " <CIUDADANO>"
                End If
                If UCase$(UserList(TempCharIndex).Name) = "GS" Then
                    Stat = Stat & " <PROGRAMADOR> ~0~245~0~1~1"
                Else
                    Stat = Stat & " <ADMINISTRADOR> ~0~205~0~1~0"
                End If
            ElseIf UserList(TempCharIndex).flags.Privilegios = 1 Then
                Stat = Stat & " <CONSEJERO> ~0~185~0~1~0"
            ElseIf UserList(TempCharIndex).flags.Privilegios = 2 Then
                Stat = Stat & " <SEMI-DIOS> ~0~185~0~1~0"
            ElseIf UserList(TempCharIndex).flags.Privilegios > 2 Then
                Stat = Stat & " <DIOS> ~0~185~0~1~0"
            ' 0.12b1
            ElseIf UserList(TempCharIndex).flags.PertAlCons > 0 Then
                    Stat = Stat & " [CONSEJO DE BANDERBILL]" & FONTTYPE_CONSEJOVesA
            ElseIf UserList(TempCharIndex).flags.PertAlConsCaos > 0 Then
                    Stat = Stat & " [CONSEJO DE LAS SOMBRAS]" & FONTTYPE_CONSEJOCAOSVesA
                    
            ElseIf Criminal(TempCharIndex) Then
                Stat = Stat & " <CRIMINAL> ~255~0~0~1~0"
            Else
                Stat = Stat & " <CIUDADANO> ~0~0~200~1~0"
            End If
            
            'If UserList(UserIndex).flags.Privilegios < 2 Then
            Call SendData(ToIndex, UserIndex, 0, Stat)
            'Else
            '    Call SendData(ToIndex, UserIndex, 0, TempCharIndex & " - " & Stat)
            'End If
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = 0
       
       End If
       
    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            
            Dim estatus As String
            
            If UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
                estatus = "(Dudoso) "
            ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) > 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 20 Then
                If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                    estatus = "(Herido) "
                Else
                    estatus = "(Sano) "
                End If
            ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) > 20 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 30 Then
                If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                    estatus = "(Malherido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                    estatus = "(Herido) "
                Else
                    estatus = "(Sano) "
                End If
            ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) > 30 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 40 Then
                If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                    estatus = "(Muy malherido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                    estatus = "(Herido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                    estatus = "(Levemente herido) "
                Else
                    estatus = "(Sano) "
                End If
            ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) > 40 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 60 Then
                If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                    estatus = "(Agonizando) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                    estatus = "(Casi muerto) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                    estatus = "(Muy Malherido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                    estatus = "(Herido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                    estatus = "(Levemente herido) "
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                    estatus = "(Sano) "
                Else
                    estatus = "(Intacto) "
                End If
            ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 60 Then
                estatus = "(" & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP) & ") "
            Else
                estatus = "!error!"
            End If

            If ConfigNPCClick <> 1 Then
                estatus = ""
            Else
                estatus = " " & estatus
            End If
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFX)
            Else
                
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    If UCase$(UserList(Npclist(TempCharIndex).MaestroUser).Name) = "GS" Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & " es mascota de ^[GS]^ [" & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP) & "]" & FONTTYPE_INFX)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & " [" & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP) & "]" & FONTTYPE_INFX)
                    End If
                Else
                    ' Mostrar Vida y Exp
                    If ConfigNPCClick = 0 Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & ". [" & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP & "] [Exp: " & CStr(Npclist(TempCharIndex).GiveEXP)) & "]" & FONTTYPE_INFX)
                    ' Mostrar Nombre nada mas
                    ElseIf ConfigNPCClick = 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & "." & FONTTYPE_INFX)
                    ' Mostrar Vida
                    ElseIf ConfigNPCClick = 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & " - Vida: " & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP) & "." & FONTTYPE_INFX)
                    ' Mostrar Vida, Exp y Oro
                    ElseIf ConfigNPCClick = 3 Then
                        Call SendData(ToIndex, UserIndex, 0, "||" & estatus & Npclist(TempCharIndex).Name & " - Vida: " & CStr(Npclist(TempCharIndex).Stats.MinHP) & "/" & CStr(Npclist(TempCharIndex).Stats.MaxHP) & " - Exp Total: " & CStr(Npclist(NpcIndex).GiveEXP) & " - ORO: " & CStr(Npclist(NpcIndex).GiveEXP) & "." & FONTTYPE_INFX)
                    End If
                End If
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFX)
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFX)
    End If
End If
Exit Sub
FalloLock:
    Call LogError("Error en LookatTitle - Fchr:" & FoundChar & " Fsme:" & FoundSomething & "N:" & Err.Number & " D:" & Err.Description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function



