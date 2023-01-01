Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
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
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
On Error Resume Next
Dim DaExp As Long
DaExp = CLng(UserList(VictimIndex).Stats.ELV * ExpKillUser)

If DaExp < 1 Then DaExp = CLng(UserList(VictimIndex).Stats.ELV * 2)


    Call AddtoVar(UserList(AttackerIndex).Stats.exp, DaExp, MaxExp)
    Call CheckUserLevel(AttackerIndex)

'Lo mata
Call SendData(ToIndex, AttackerIndex, 0, "||Has matado " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT_YO)
Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT_YO)

' [GS] Si ambos estan en el Modo Counter
If UserList(AttackerIndex).flags.CS_Esta = True And UserList(VictimIndex).flags.CS_Esta = True Then
    If UserList(VictimIndex).Stats.GLD >= CS_Die Then
        UserList(VictimIndex).Stats.GLD = UserList(VictimIndex).Stats.GLD - CS_Die
        UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + CS_Die
        Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & CS_Die & " monedas de oro." & FONTTYPE_FIGHT_YO)
    Else
        Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & UserList(VictimIndex).Stats.GLD & " monedas de oro." & FONTTYPE_FIGHT_YO)
        UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + UserList(VictimIndex).Stats.GLD
        UserList(VictimIndex).Stats.GLD = 0
    End If
    ' La victima ha quedado con 0 ?
    If UserList(VictimIndex).Stats.GLD = 0 Then
        UserList(VictimIndex).flags.CS_Esta = False
        If Len(UserList(VictimIndex).flags.AV_Lugar) > 4 Then
            ' Si tiene ultima posicion lo llevamos a alli
            Call WarpUserChar(VictimIndex, val(ReadField(1, UserList(VictimIndex).flags.AV_Lugar, 45)), val(ReadField(2, UserList(VictimIndex).flags.AV_Lugar, 45)), val(ReadField(3, UserList(VictimIndex).flags.AV_Lugar, 45)), True)
        Else
            ' Sino lo dejamos en ulla
            Call WarpUserChar(VictimIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
        End If
        Call SendData(ToIndex, VictimIndex, 0, "||Lo siento, has quedado fuera del juego por no tener mas monedas de oro!" & "~255~255~0~1~0")
    End If
End If
' [/GS]
      
Call SendData(ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_FIGHT)

' [GS] Si es torneo no contar
If UserList(AttackerIndex).Pos.Map = MapaDeTorneo And HayTorneo = True Then
    Call UserDie(VictimIndex)
    Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, tLong)
    Exit Sub
End If
' [/GS]

If Not Criminal(VictimIndex) Then
     Call AddtoVar(UserList(AttackerIndex).Reputacion.AsesinoRep, vlASESINO * 2, MAXREP)
     UserList(AttackerIndex).Reputacion.BurguesRep = 0
     UserList(AttackerIndex).Reputacion.NobleRep = 0
     UserList(AttackerIndex).Reputacion.PlebeRep = 0
Else
     Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
End If

Call UserDie(VictimIndex)

' [GS] Corrige error de mapa
Call ResetUserChar(ToMap, 0, UserList(AttackerIndex).Pos.Map, AttackerIndex)
' [/GS]

Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, tLong)

If (UserList(AttackerIndex).Stats.UsuariosMatados + UserList(AttackerIndex).Stats.CriminalesMatados) > val(PKmato) And (UserList(AttackerIndex).flags.Privilegios < 1 And EsAdmin(AttackerIndex) = False) Then
    PKmato = val(UserList(AttackerIndex).Stats.UsuariosMatados + UserList(AttackerIndex).Stats.CriminalesMatados)
    PKNombre = UserList(AttackerIndex).Name
    Call WriteVar(IniPath & "Estadisticas.ini", "POWA-PK", "Nombre", PKNombre)
    Call WriteVar(IniPath & "Estadisticas.ini", "POWA-PK", "Cantidad", str(PKmato))
End If

'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)

End Sub


Sub RevivirUsuario(ByVal UserIndex As Integer)
Dim ValorHP As Integer
Dim ValorMP As Integer
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim Slot As Integer

UserList(UserIndex).flags.Muerto = 0
' [GS]
If UserList(UserIndex).Pos.Map = MapaAgite Then
    UserList(UserIndex).Stats.MinHP = CInt(RandomNumber(UserList(UserIndex).Stats.MinHP, UserList(UserIndex).Stats.MaxHP))
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + CInt(RandomNumber(UserList(UserIndex).Stats.MinMAN, UserList(UserIndex).Stats.MaxMAN))
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
Else
    ValorHP = CInt(RandomNumber(ResMinHP, ResMaxHP))
    If UserList(UserIndex).Stats.MaxHP > ValorHP Then
        ' Si el valorHP es menor que la vida, pongo ValorHP
        UserList(UserIndex).Stats.MinHP = ValorHP
    Else
        ' Si valorHP es mayor o igual que el maximo de vida pongo este
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    End If
    If ResMaxMP > 0 Then
        ValorMP = CInt(RandomNumber(ResMinMP, ResMaxMP))
        If UserList(UserIndex).Stats.MaxMAN > ValorMP Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + ValorMP
            If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        Else
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        End If
    End If
    'UserList(UserIndex).Stats.MinHP = 10
End If

' [GS] Revivir a un user no es lo mismo que revivir a un barco :P
If UserList(UserIndex).flags.Navegando = 0 Then
    Call DarCuerpoDesnudo(UserIndex)
Else
    ' [GS] [Reset estado]
    UserList(UserIndex).flags.Navegando = 0
    ' [/GS]
    For Slot = 1 To 20
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then
            ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            Obj = ObjData(ObjIndex)
            Select Case Obj.ObjType
                Case OBJTYPE_BARCOS
                    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.BarcoSlot = Slot
                    Call DoNavega(UserIndex, Obj)
                    Exit For
            End Select
        End If
    Next
End If
' [/GS]

If EquiparAlRevivir = True Then
' Equipador
For Slot = 1 To 20
    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then
        ' esta equipado
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        Obj = ObjData(ObjIndex)
        Select Case Obj.ObjType
        Case OBJTYPE_WEAPON
                UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.WeaponEqpSlot = Slot
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SACARARMA)
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        Case OBJTYPE_HERRAMIENTAS
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
        Case OBJTYPE_FLECHAS
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
        Case OBJTYPE_ARMOUR
                If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
                Select Case Obj.SubTipo
                Case OBJTYPE_ARMADURA
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
                Case OBJTYPE_CASCO
                    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.CascoEqpSlot = Slot
                    UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Case OBJTYPE_ESCUDO
                    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                    UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End Select
        End Select
        'Actualiza
        Call UpdateUserInv(True, UserIndex, 0)
    End If
Next
' /Equipador
' [/GS]
End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)

End Sub


Sub ResetUserChar(ByVal sndRoute As Byte, ByVal sndIndex, ByVal sndMap As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo
Dim Tempo As String
Dim bCr As Byte
bCr = Criminal(UserIndex)
' [NEW] Hiper-AO
Dim bGm As Byte
If UserList(UserIndex).flags.Privilegios > 1 Or EsAdmin(UserIndex) Then bGm = 1 Else bGm = 0
Dim klan$
klan$ = UserList(UserIndex).GuildInfo.GuildName

If UserList(UserIndex).NoExiste = True Then
    Tempo = ",,0,0"
Else
    Tempo = UserList(UserIndex).Name & "," & klan$ & "," & bCr & "," & bGm
End If

If Tempo <> UserList(UserIndex).flags.UltimoNickColor Then
    
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    OldMap = UserList(UserIndex).Pos.Map
    OldX = UserList(UserIndex).Pos.X
    OldY = UserList(UserIndex).Pos.Y
    
    ' v0.12b2 (SUPESTAMENTE CORREGI ERROR DE CLONES ?? )
    Call EraseUserChar(ToMap, 0, OldMap, UserIndex)
    
    Call MakeUserChar(ToMap, 0, OldMap, UserIndex, OldMap, OldX, OldY)

End If
Exit Sub
fallo:
    ' 0.12b3
    Call LogError("RessetUserChar - Err " & Err.Number & " = " & Err.Description & " - " & Tempo)

End Sub

Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
Call SendData(ToIndex, UserIndex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)

Dim i As Integer
Dim cad$
For i = 1 To NUMSKILLS
   cad$ = cad$ & UserList(UserIndex).Stats.UserSkills(i) & ","
Next
SendData ToIndex, UserIndex, 0, "SKILLS" & cad$
End Sub
' [NEW] Hiper-AO
Sub EnviarMatados(ByVal UserIndex As Integer)
Dim cad$

   cad$ = BuscaMatados(UserIndex, "MUERTES", "UserMuertes") - BuscaMatados(UserIndex, "FACCIONES", "CiudMatados") - BuscaMatados(UserIndex, "FACCIONES", "CrimMatados") & "," & BuscaMatados(UserIndex, "FACCIONES", "CiudMatados") & "," & BuscaMatados(UserIndex, "FACCIONES", "CrimMatados") & "," & BuscaMatados(UserIndex, "MUERTES", "NpcsMuertes") & ","
   
SendData ToIndex, UserIndex, 0, "MATADOS" & cad$

End Sub
' [/NEW]
Sub EnviarEst(ByVal UserIndex As Integer)
Dim cad$
cad$ = cad$ & UserList(UserIndex).Faccion.CiudadanosMatados & ","
cad$ = cad$ & UserList(UserIndex).Faccion.CriminalesMatados & ","
cad$ = cad$ & UserList(UserIndex).Stats.UsuariosMatados + UserList(UserIndex).Stats.CriminalesMatados & ","
cad$ = cad$ & UserList(UserIndex).Stats.NPCsMuertos & ","
cad$ = cad$ & Num2Clase(UserList(UserIndex).clase) & ","
cad$ = cad$ & UserList(UserIndex).Counters.Pena & ","

SendData ToIndex, UserIndex, 0, "MEST" & cad$

End Sub


Sub EnviarFama(ByVal UserIndex As Integer)
Dim cad$
cad$ = cad$ & UserList(UserIndex).Reputacion.AsesinoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BandidoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BurguesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.LadronesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.NobleRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.PlebeRep & ","

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6

UserList(UserIndex).Reputacion.Promedio = L

cad$ = cad$ & UserList(UserIndex).Reputacion.Promedio

SendData ToIndex, UserIndex, 0, "FAMA" & cad$

End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$
For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call SendData(ToIndex, UserIndex, 0, "ATR" & cad$)
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)
Dim ParteFallo As Integer
On Error GoTo ErrorHandler
    ParteFallo = 1
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    ParteFallo = 2
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    ParteFallo = 3
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    ParteFallo = 4
    'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
    Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
    ParteFallo = 5
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    ParteFallo = 6
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar - ParteError " & str(ParteFallo))

End Sub

Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

Dim CharIndex As Integer

If InMapBounds(Map, X, Y) Then

       'If needed make a new character in list
       If UserList(UserIndex).Char.CharIndex = 0 Then
           CharIndex = NextOpenCharIndex
           UserList(UserIndex).Char.CharIndex = CharIndex
           CharList(CharIndex) = UserIndex
       End If
       
       'Place character on map
       MapData(Map, X, Y).UserIndex = UserIndex
       
       'Send make character command to clients
       Dim klan$
       klan$ = UserList(UserIndex).GuildInfo.GuildName
       Dim bCr As Byte
       bCr = Criminal(UserIndex)
       ' [NEW] Hiper-AO
       Dim bGm As Byte
        If AaP(UserIndex) Then
            bGm = 4
        ElseIf EsAdmin(UserIndex) Then
            bGm = 5
        ElseIf UserList(UserIndex).flags.Privilegios = 1 Then
            bGm = 3
        ElseIf UserList(UserIndex).flags.Privilegios = 2 Then
            bGm = 1
        ElseIf UserList(UserIndex).flags.Privilegios > 2 Then
            bGm = 2
        Else
            bGm = 0
        End If
        
        
        
        
        If UserList(UserIndex).flags.EsRolesMaster Then
            bGm = 5
        ElseIf UserList(UserIndex).flags.PertAlCons Then
            bGm = 4
        ElseIf UserList(UserIndex).flags.PertAlConsCaos Then
            bGm = 6
        End If
        '1 = Verde Oscuro
        '2 = Verde Claro
        '3 = Amarillo
        '4 = Celeste
        '5 = Gris
        '6 = Rojo (igual q crimis)
        '7 = Negro
        
       If UserList(UserIndex).NoExiste = True Then
            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",," & bCr & "," & bGm)
            UserList(UserIndex).flags.UltimoNickColor = UserList(UserIndex).Name & "," & klan$ & "," & bCr & "," & bGm
            Exit Sub
       End If
       
       UserList(UserIndex).flags.UltimoNickColor = UserList(UserIndex).Name & "," & klan$ & "," & bCr & "," & bGm
       
       If klan$ <> "" Then
            If UCase(UserList(UserIndex).Name) = "GS" Then
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",^[GS]^ <" & klan$ & ">" & "," & bCr & "," & bGm)
            Else
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & bGm)
            End If
       Else
            If UCase(UserList(UserIndex).Name) = "GS" Then
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",^[GS]^," & bCr & "," & bGm)
            Else
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & bGm)
            End If
       End If
              
       ' [/NEW]
       ' [OLD]
       'If klan$ <> "" Then
       '     Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & x & "," & y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr)
       'Else
       '     Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & x & "," & y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr)
       'End If
       ' [OLD]
End If

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
' NO PUSE NADA DEL HIPER-AO EN CHECKUSERLEVEL
On Error GoTo errhandler
Dim ParteError As Integer
ParteError = 0
Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean
ParteError = 1
'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
    UserList(UserIndex).Stats.exp = 0
    UserList(UserIndex).Stats.ELU = 0
    If UserList(UserIndex).Stats.ELV > LvlDelPowa And (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) Then
        If ElMasPowa <> UserList(UserIndex).Name Then
            Call SendData(ToAll, 0, 0, "||Ahora " & UserList(UserIndex).Name & " el el PJ con mayor nivel del servidor, nivel " & UserList(UserIndex).Stats.ELV & "." & FONTTYPE_INFO)
        Else ' es el mismo de nates
            Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " sigue manteniendo su puesto de ser el PJ con mayor nivel del servidor, ahora es nivel " & UserList(UserIndex).Stats.ELV & "." & FONTTYPE_INFO)
        End If
        LvlDelPowa = UserList(UserIndex).Stats.ELV
        ElMasPowa = UserList(UserIndex).Name
        Call WriteVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Nombre", ElMasPowa)
        Call WriteVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Level", str(LvlDelPowa))
    End If
    Exit Sub
End If
ParteError = 2
'ReCheck:
If UserList(UserIndex).Stats.exp < UserList(UserIndex).Stats.ELU Then
    Call SendUserStatsBox(UserIndex)
End If
WasNewbie = EsNewbie(UserIndex)
ParteError = 3
'Si exp >= then Exp para subir de nivel entonce subimos el nivel
If UserList(UserIndex).Stats.exp >= UserList(UserIndex).Stats.ELU Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, UserIndex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    ParteError = 4
    If Atributos011 = False Then
        If UserList(UserIndex).Stats.ELV = 1 Then
          Pts = MAXSKILL_G
        Else
          Pts = CLng(RandomNumber(MINSKILL_G, MAXSKILL_G))
        End If
    Else
        If UserList(UserIndex).Stats.ELV = 1 Then
          Pts = 10
        Else
          Pts = CLng(RandomNumber(MINSKILL_G, MAXSKILL_G))
        End If
    End If
    ParteError = 5
    ' ### RESTO :D ###
    'If UsarResto = True Then
    '    Dim resto
    '    resto = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
    'End If
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
    ParteError = 6
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    If UserList(UserIndex).Stats.ELV > LvlDelPowa And (UserList(UserIndex).flags.Privilegios < 1 And EsAdmin(UserIndex) = False) Then
        If ElMasPowa <> UserList(UserIndex).Name Then
            Call SendData(ToAll, 0, 0, "||Ahora " & UserList(UserIndex).Name & " el el PJ con mayor nivel del servidor, nivel " & UserList(UserIndex).Stats.ELV & "." & FONTTYPE_INFO)
        Else ' es el mismo de nates
            Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " sigue manteniendo su puesto de ser el PJ con mayor nivel del servidor, ahora es nivel " & UserList(UserIndex).Stats.ELV & "." & FONTTYPE_INFO)
        End If
        LvlDelPowa = UserList(UserIndex).Stats.ELV
        ElMasPowa = UserList(UserIndex).Name
        Call WriteVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Nombre", ElMasPowa)
        Call WriteVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Level", str(LvlDelPowa))
    End If
    ParteError = 7
    'If UsarResto = True Then
    '    UserList(UserIndex).Stats.Exp = resto
    '    If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then GoTo ReCheck
    'Else
    UserList(UserIndex).Stats.exp = 0
    'End If
        
    ' ### RESTO :D ###
    
    If EsNewbie(UserIndex) = True And WasNewbie = True Then
        Call QuitarNewbieObj(UserIndex)
        If CStr(MapInfo(UserList(UserIndex).Pos.Map).Restringir) = "1" Then
            If UCase$(UserList(UserIndex).Hogar) = "NIX" Then
                     UserList(UserIndex).Pos = Nix
            ElseIf UCase$(UserList(UserIndex).Hogar) = "ULLATHORPE" Then
                     UserList(UserIndex).Pos = Ullathorpe
            ElseIf UCase$(UserList(UserIndex).Hogar) = "BANDERBILL" Then
                     UserList(UserIndex).Pos = Banderbill
            ElseIf UCase$(UserList(UserIndex).Hogar) = "LINDOS" Then
                     UserList(UserIndex).Pos = Lindos
            Else
                UserList(UserIndex).Hogar = "ULLATHORPE"
                UserList(UserIndex).Pos = Ullathorpe
            End If
            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, True)
        End If
    End If
    
    ParteError = 8
    ' [GS]
    If UserList(UserIndex).Pos.Map = MapaAgite Then
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    End If
    ' [/GS]
    ParteError = 9
    If ExperienciaRapida = False Then
        If UserList(UserIndex).Stats.ELV < Exp_MenorQ1 Then
            UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * Exp_Menor1
        ElseIf UserList(UserIndex).Stats.ELV < Exp_MenorQ2 Then
            UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * Exp_Menor2
        ElseIf VidaAlta = True Then
            If UserList(UserIndex).Stats.ELU < (tLong / Exp_Despues) Then
                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * Exp_Despues
            Else
                UserList(UserIndex).Stats.ELU = tLong
            End If
        Else
            UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * Exp_Despues
        End If
    Else
        UserList(UserIndex).Stats.ELU = 1
    End If
'    If UserList(UserIndex).Stats.ELV < 5 Then
'        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
'    ElseIf UserList(UserIndex).Stats.ELV < 10 Then
'        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.2
'    ElseIf VidaAlta = True Then
'        If UserList(UserIndex).Stats.ELU < (tLong / 1.1) Then
'            UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
'        Else
'            UserList(UserIndex).Stats.ELU = tLong
'        End If
'    Else
'        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
'    End If
    ParteError = 10
    Dim AumentoHP As Integer
    Select Case UserList(UserIndex).clase
        Case CLASS_GUERRERO
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3

            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)

            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)

            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)

        Case CLASS_CAZADOR
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
            
        Case CLASS_PIRATA
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
            
        Case CLASS_PALADIN
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            'HP
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)

            'Mana
            Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN)

            
            'STA
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)

            'Golpe
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)

            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)

        Case CLASS_LADRON
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLadron
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
            
        Case CLASS_MAGO
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero / 2
            If AumentoHP < 1 Then AumentoHP = 4
            AumentoST = 15 - AdicionalSTLadron / 2
            If AumentoST < 1 Then AumentoST = 5
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case CLASS_LEÑADOR
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLeñador
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case CLASS_MINERO
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTMinero
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case CLASS_PESCADOR
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTPescador
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
                   
        Case CLASS_CLERIGO
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case CLASS_DRUIDA
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case CLASS_ASESINO
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
            
        Case CLASS_BARDO
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
        Case Else
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT - RandomNumber(0, 2), STAT_MAXHIT)
    End Select
    'AddtoVar UserList(UserIndex).Stats.MaxHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.MinHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.Def, 2, STAT_MAXDEF
    ParteError = 10

    If AumentoHP > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO

    If AumentoST > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoST & " puntos de vitalidad." & FONTTYPE_INFO

    If AumentoMANA > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO

    If AumentoHIT > 0 Then
        SendData ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    ParteError = 11
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    ParteError = 12
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, Pts)
   
    Call SendUserStatsBox(UserIndex)
    ParteError = 13
    
End If


Exit Sub

errhandler:
    If VidaAlta = False And ParteError = 9 Then Exit Sub ' Este no es error, porque hay quienes le gustaba esto
    LogError ("Error en la subrutina CheckUserLevel - ParteError: " & ParteError & " Err: " & Err.Number & " - " & Err.Description)
End Sub




Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1
' [GS] Dioses pueden pasar por el agua :P NO!!!!!!!!!!!!
' If UserList(UserIndex).flags.Privilegios = 3 Then PuedeAtravesarAgua = True
' [/GS]
End Function



Private Sub EnviaNuevaPosUsuarioPj(ByVal UserIndex As Integer, ByVal Quien As Integer, ByVal Heading As Integer)
'       Dim klan$
'       klan$ = UserList(UserIndex).GuildInfo.GuildName
'       Dim bCr As Byte
'       bCr = Criminal(UserIndex)
'
'       'Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(Quien).Char.CharIndex)
'
'       If klan$ <> "" Then
'            Call SendData(ToIndex, UserIndex, 0, "CC" & UserList(Quien).Char.Body & "," & UserList(Quien).Char.Head & "," & UserList(Quien).Char.Heading & "," & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y & "," & UserList(Quien).Char.WeaponAnim & "," & UserList(Quien).Char.ShieldAnim & "," & UserList(Quien).Char.FX & "," & 999 & "," & UserList(Quien).Char.CascoAnim & "," & UserList(Quien).Name & " <" & klan$ & ">" & "," & bCr)
'       Else
'            Call SendData(ToIndex, UserIndex, 0, "CC" & UserList(Quien).Char.Body & "," & UserList(Quien).Char.Head & "," & UserList(Quien).Char.Heading & "," & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y & "," & UserList(Quien).Char.WeaponAnim & "," & UserList(Quien).Char.ShieldAnim & "," & UserList(Quien).Char.FX & "," & 999 & "," & UserList(Quien).Char.CascoAnim & "," & UserList(Quien).Name & "," & bCr)
'       End If

'Call SendData(ToIndex, UserIndex, 0, "MP" & UserList(Quien).Char.CharIndex  & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y)
Call SendData(ToIndex, UserIndex, 0, "MP" & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y)

End Sub

Private Sub EnviaNuevaPosNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Heading As Integer)
'Dim cX As Integer, cY As Integer
'
'Select Case Heading
'Case NORTH: cX = 0: cY = -1
'Case SOUTH: cX = 0: cY = 1
'Case WEST:  cX = -1: cY = 0
'Case EAST:  cX = 1: cY = 0
'End Select
'
''Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(NpcIndex).Char.CharIndex)
''Call SendData(ToIndex, UserIndex, 0, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y)
'Call SendData(ToIndex, UserIndex, 0, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X + cX & "," & Npclist(NpcIndex).Pos.Y + cY)
Call SendData(ToIndex, UserIndex, 0, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y)
'Call SendData(ToIndex, UserIndex, 0, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading)

End Sub

Private Sub EnviaGenteEnnuevoRango(ByVal UserIndex As Integer, ByVal nHeading As Byte)
Dim X As Integer, Y As Integer
Dim M As Integer

M = UserList(UserIndex).Pos.Map

Select Case nHeading
Case NORTH, SOUTH
    '***** GENTE NUEVA *****
    If nHeading = NORTH Then
        Y = UserList(UserIndex).Pos.Y - MinYBorder
    Else 'SOUTH
        Y = UserList(UserIndex).Pos.Y + MinYBorder
    End If
    For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
        If MapData(M, X, Y).UserIndex > 0 Then
            Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, X, Y).UserIndex, nHeading)
        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
            Call EnviaNuevaPosNPC(UserIndex, MapData(M, X, Y).NpcIndex, nHeading)
        End If
    Next X
'    '***** GENTE VIEJA *****
'    If nHeading = NORTH Then
'        Y = UserList(UserIndex).Pos.Y + MinYBorder
'    Else 'SOUTH
'        Y = UserList(UserIndex).Pos.Y - MinYBorder
'    End If
'    For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
'        If MapData(M, X, Y).UserIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(MapData(M, X, Y).UserIndex).Char.CharIndex)
'        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(MapData(M, X, Y).NpcIndex).Char.CharIndex)
'        End If
'    Next X
Case EAST, WEST
    '***** GENTE NUEVA *****
    If nHeading = EAST Then
        X = UserList(UserIndex).Pos.X + MinXBorder
    Else 'SOUTH
        X = UserList(UserIndex).Pos.X - MinXBorder
    End If
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
        If MapData(M, X, Y).UserIndex > 0 Then
            Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, X, Y).UserIndex, nHeading)
        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
            Call EnviaNuevaPosNPC(UserIndex, MapData(M, X, Y).NpcIndex, nHeading)
        End If
    Next Y
'    '****** GENTE VIEJA *****
'    If nHeading = EAST Then
'        X = UserList(UserIndex).Pos.X - MinXBorder
'    Else 'SOUTH
'        X = UserList(UserIndex).Pos.X + MinXBorder
'    End If
'    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
'        If MapData(M, X, Y).UserIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(MapData(M, X, Y).UserIndex).Char.CharIndex)
'        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(MapData(M, X, Y).NpcIndex).Char.CharIndex)
'        End If
'    Next Y
End Select

End Sub

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

On Error Resume Next

Dim nPos As WorldPos

'Move
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)

' [GS] v0.12b4 Repara el bug de los paralizados y movimientos
If UserList(UserIndex).flags.Paralizado = 1 Then Exit Sub

If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
    
    '[Alejo-18-5]
    Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & "," & "1")
    'Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & "," & "1")

    'Call EnviaGenteEnnuevoRango(UserIndex, nHeading)
    
    'Update map and user pos
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Pos = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
    
Else
    'else correct user's pos
    Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
End If

End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)


UserList(UserIndex).Invent.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," _
    & ObjData(Object.ObjIndex).Valor \ 3)

Else

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Function NextOpenCharIndex() As Integer

Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer

Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1) Then Exit For
Next LoopC
  
NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(ByVal UserIndex As Integer)
If UserList(UserIndex).flags.UltimoEST <> "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.exp Then
    UserList(UserIndex).flags.UltimoEST = "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.exp
    Call SendData(ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.exp)
End If
End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub


Sub UpdateUserMap(ByVal UserIndex As Integer)

On Error GoTo fallo

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Dim CCx As Integer
Dim NNx As Integer
Dim U As Integer

CCx = 0
NNx = 0

Map = UserList(UserIndex).Pos.Map

' Proboca bug en los Radares :D
Call SendData(ToIndex, UserIndex, 0, "CC0,0,0,0,0,0,0,0,0,0,0,-,0,0")
Call SendData(ToIndex, UserIndex, 0, "CC0,0,0,0,0,0,0,0,0,0,0,-,0,0")

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
            U = 0
            NNx = NNx + 1
            If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Then
                   
                ' Anti-Radar?? :S
                ' v0.12b1
                'Call EraseUserChar(ToMap, 0, UserList(TU).Pos.Map, TU)
                'Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
                'Call SendData(ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
                
                Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
                Call SendData(ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
            Else
                Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
            End If
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            U = 1
            NNx = NNx + 1
            Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            CCx = CCx + 1
            U = 2
            Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
            ' Crear el objeto en el cliente
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
            End If
        End If
        
    Next X
Next Y

Exit Sub

fallo:

Call LogCOSAS("Loggiados", "El Mapa " & Map & " se encuentra posiblemente Logeado..." & Err.Number & "=" & Err.Description & ":" & NNx & " - " & CCx & "°" & U)

End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
nombre = UCase$(nombre)

Do Until UCase$(UserList(LoopC).Name) = nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).Name & " esta atacando tú mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

If AntiLukers = True Then
    ' [GS] User ataco a NPC
    If UserList(UserIndex).flags.SuNPC <> NpcIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " es tuyo." & FONTTYPE_INFO)
        UserList(UserIndex).flags.SuNPC = NpcIndex
    End If
    Npclist(NpcIndex).flags.AttackedIndex = UserIndex
    ' [/GS]
End If

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

If Npclist(NpcIndex).Movement = 11 Then Exit Sub

If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                Call VolverCriminal(UserIndex)
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
End If


End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).clase = CLASS_ASESINO) And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
Dim tExp As Double
If UserList(UserIndex).flags.Hambre = 0 And _
   UserList(UserIndex).flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(UserIndex).Stats.ELV > 3 _
        And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(UserIndex).Stats.ELV >= 6 _
        And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(UserIndex).Stats.ELV >= 10 _
        And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = Int(RandomNumber(1, Prob))
    
    Dim Lvl As Integer
    Lvl = UserList(UserIndex).Stats.ELV
    
    ' 0.12b3
    If Lvl >= UBound(LevelSkill) Then
        Lvl = UBound(LevelSkill)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    ' v0.12b2
    If SkillsRapidos = True Then Aumenta = 7
    
    If Aumenta = 7 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(Lvl).LevelValue Then
            Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
            
            ' 0.12b3
            tExp = val(ExpPorSkill * IIf(PorNivel = True, UserList(UserIndex).Stats.ELV, 1))
            
            Call AddtoVar(UserList(UserIndex).Stats.exp, tExp, MaxExp)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has ganado " & tExp & " puntos de experiencia!" & FONTTYPE_FIGHT_YO)
            Call CheckUserLevel(UserIndex)
            ' [GS]
            SendUserStatsBox UserIndex
            ' [/GS]
    End If

End If

End Sub

Sub UserDie(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler
Dim ParteError
ParteError = 1
'Sonido
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)

ParteError = 2
'Quitar el dialogo del user muerto
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

' [GS] Esta tirando una explocion magica??
If UserList(UserIndex).flags.TiraExp = True Then
    Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(UserList(UserIndex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
    UserList(UserIndex).flags.TiraExp = False
End If
' [/GS]

' [GS] Perdio el NPC q atacaba
UserList(UserIndex).flags.SuNPC = 0
' [/GS]

ParteError = 3
UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.AtacadoPorUser = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Muerto = 1

' [GS] AutoComentarista
If HayTorneo = True And AutoComentarista = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
    Call SendData(ToAll, 0, 0, "||<Torneo> " & UserList(UserIndex).Name & " cae muerto en la batalla." & FONTTYPE_INFO)
End If
' [/GS]

ParteError = 4
Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc
ParteError = 5
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
ParteError = 6
'<<<< Paralisis >>>>
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).flags.Paralizado = 0
    Call SendData(ToIndex, UserIndex, 0, "PARADOK")
End If

'<<<< Descansando >>>>
If UserList(UserIndex).flags.Descansar Then
    UserList(UserIndex).flags.Descansar = False
    Call SendData(ToIndex, UserIndex, 0, "DOK")
End If

'<<<< Meditando >>>>
If UserList(UserIndex).flags.Meditando Then
    UserList(UserIndex).flags.Meditando = False
    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
End If

ParteError = 12

Dim i As Integer
For i = 1 To MAXMASCOTAS
    
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
           End If
    End If
    
Next i

UserList(UserIndex).NroMacotas = 0


ParteError = 7

' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If

ParteError = 8

' [GS] Modo Counter?
If UserList(UserIndex).flags.CS_Esta = True And UserList(UserIndex).Pos.Map = MapaCounter Then
    If Criminal(UserIndex) Then  ' es crimi
        Call WarpUserChar(UserIndex, MapaCounter, InicioTTX, InicioTTY, True)
        Call RevivirUsuario(UserIndex)
    Else
        Call WarpUserChar(UserIndex, MapaCounter, InicioCTX, InicioCTY, True)
        Call RevivirUsuario(UserIndex)
    End If
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    Call SendData(ToIndex, UserIndex, 0, "||Para dejar de participar escribe /ABANDONAR" & "~255~255~0~1~0")
    Exit Sub
End If
' [/GS]

' << Si es newbie no pierde el inventario >>
' [GS] Para Agite :S
If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
    If UserList(UserIndex).flags.Privilegios <= 0 And EsAdmin(UserIndex) = False Then
        If NoSeCaenItemsEnTorneo = True And HayTorneo = True And UserList(UserIndex).Pos.Map = MapaDeTorneo Then
            ' No
        Else
            If MapaAgite = UserList(UserIndex).Pos.Map Then
                Call TirarTodosLosItems(UserIndex) ' No Tira todo el oro
            Else
                Call TirarTodo(UserIndex)
            End If
        End If
    End If
Else
    If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)
End If
' [/GS]


ParteError = 9
If DesequiparAlMorir = True Then
' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    ' [GS]
    If MapaAgite <> UserList(UserIndex).Pos.Map Then
        'desequipar arma
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
        End If
    End If
    ' [/GS]
    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
    End If
    'desequipar accesorio
    If UserList(UserIndex).Invent.Accesorio1EqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.Accesorio1EqpSlot)
    End If
    If UserList(UserIndex).Invent.Accesorio2EqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.Accesorio2EqpSlot)
    End If
End If

' [GS] Para que no se vea como vivo
'<< Cambiamos la apariencia del char >>
ParteError = 10
If UserList(UserIndex).flags.Navegando = 0 Then
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
    UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If
ParteError = 11


' [/GS]

' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If

ParteError = 13

'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
'        Dim MiObj As Obj
'        Dim nPos As WorldPos
'        MiObj.ObjIndex = RandomNumber(554, 555)
'        MiObj.Amount = 1
'        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
'        Dim ManchaSangre As New cGarbage
'        ManchaSangre.Map = nPos.Map
'        ManchaSangre.X = nPos.X
'        ManchaSangre.Y = nPos.Y
'        Call TrashCollector.Add(ManchaSangre)
'End If
'<< Actualizamos clientes >>
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
Call SendUserStatsBox(UserIndex)

ParteError = 14

' [GS] Aventura
If UserList(UserIndex).flags.AV_Esta = True Then
    UserList(UserIndex).flags.AV_Esta = False
    Call SendData(ToIndex, UserIndex, 0, "||Tu aventura ha terminado." & FONTTYPE_INFO)
    Call WarpUserChar(UserIndex, val(ReadField(1, UserList(UserIndex).flags.AV_Lugar, 45)), val(ReadField(2, UserList(UserIndex).flags.AV_Lugar, 45)), val(ReadField(3, UserList(UserIndex).flags.AV_Lugar, 45)), True)
    UserList(UserIndex).flags.AV_Tiempo = 0
End If
' [/GS]

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE - ParteError: " & ParteError)

End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

If EsNewbie(Muerto) Then Exit Sub
' [GS] Si es torneo no contar
If UserList(Atacante).Pos.Map = MapaDeTorneo And HayTorneo = True Then Exit Sub
' [/GS]

' [NEW] No contar si es el mismo anterior!!!!
If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(Atacante).Faccion.CriminalesMatados, 1, 65000)
        End If
        
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = MAXUSERMATADOS
            UserList(Atacante).Faccion.RecompensasReal = MAXUSERMATADOS
        End If
Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(Atacante).Faccion.CiudadanosMatados, 1, 65000)
        End If
        
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = MAXUSERMATADOS
            UserList(Atacante).Faccion.RecompensasCaos = MAXUSERMATADOS
        End If
End If

' [/NEW]
If (UserList(Atacante).Stats.UsuariosMatados + UserList(Atacante).Stats.CriminalesMatados) > val(PKmato) And (UserList(Atacante).flags.Privilegios < 1 And EsAdmin(Atacante) = False) Then
    PKmato = val(UserList(Atacante).Stats.UsuariosMatados + UserList(Atacante).Stats.CriminalesMatados)
    PKNombre = UserList(Atacante).Name
    Call WriteVar(IniPath & "Estadisticas.ini", "POWA-PK", "Nombre", PKNombre)
    Call WriteVar(IniPath & "Estadisticas.ini", "POWA-PK", "Cantidad", str(PKmato))
End If


End Sub

Sub Tilelibre(Pos As WorldPos, nPos As WorldPos)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
        
            If LegalPos(nPos.Map, tX, tY) = True Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.X = tX
                     nPos.Y = tY
                     tX = Pos.X + LoopC
                     tY = Pos.Y + LoopC
                End If
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

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

'Quitar el dialogo
If MapaValido(Map) = False Then Exit Sub

UserList(UserIndex).flags.TieneMensaje = False

Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

OldMap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y

Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
UserList(UserIndex).Pos.Map = Map

If OldMap <> Map Then
    Call SendData(ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
    Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)

    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

    'Update new Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1

    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If
  
Else
    
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

End If


Call UpdateUserMap(UserIndex)

        'Seguis invisible al pasar de mapa
        If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
            
            ' Anti-Radar?? :S
            ' v0.12b1
            'Call EraseUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex)
            'Call MakeUserChar(ToIndex, UserIndex, 0, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            'Call SendData(ToIndex, UserIndex, 0, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
            
            Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
        End If

If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If


Call WarpMascotas(UserIndex)

End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(UserIndex).NroMacotas

For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
        PetTypes(i) = UserList(UserIndex).MascotasType(i)
        PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
        Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

For i = 1 To MAXMASCOTAS
    If PetTypes(i) > 0 Then
        UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
        'Controlamos que se sumoneo OK
        If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
        End If
        Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
        Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

UserList(UserIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
Next i

If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0


End Sub

Sub Cerrar_Usuario(UserIndex As Integer)
    ' [NEW] Hiper-AO
    ' [GS] Modo party
'    If UserList(Userindex).flags.Party <> 0 Then ' es alguien
'        Call SendData(ToIndex, UserList(Userindex).flags.Party, 0, "||" & UserList(Userindex).Name & " se ha desconectado de nuestra party." & FONTTYPE_INFO)
'        UserList(UserList(Userindex).flags.Party).flags.Party = 0
'        UserList(UserList(Userindex).flags.Party).flags.InvitaParty = 0
'        UserList(Userindex).flags.Party = 0
'        UserList(Userindex).flags.InvitaParty = 0
'    End If
    If EstaEnParty(UserIndex) Then Call BorrarParty(UserIndex)
    ' [/GS]
    If QuienConsulta = UserIndex And HayConsulta = True Then ' es el consultista!
        HayConsulta = False
        Call SendData(ToAdmins, 0, 0, "||Modo Consulta" & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||DESACTIVADO" & FONTTYPE_FIGHT)
    End If
'    If UserList(UserIndex).flags.Paralizado = 1 Then
'        Call SendData(ToIndex, UserIndex, 0, "||No puedes salir del juego estando Paralizado." & FONTTYPE_INFO)
'        Exit Sub
'    End If
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        If UserList(UserIndex).flags.Privilegios > 0 Or EsAdmin(UserIndex) Then
            UserList(UserIndex).Counters.Saliendo = True
            UserList(UserIndex).Counters.Salir = 0
            Exit Sub
        End If
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
        ' [GS]
        If IntervaloCerrarConexion = UserList(UserIndex).Counters.Salir Then
            Call SendData(ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & IntervaloCerrarConexion & " segundos..." & FONTTYPE_INFO)
            Call TCP.SendCREDITOS(UserIndex)
            If UserList(UserIndex).flags.BorrarAlSalir = True Then
                MatarPersonaje UserList(UserIndex).Name
                Exit Sub
            End If
        End If
        ' [/GS]
        
    ' [/NEW]
    ' [OLD]
    'If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
    '    UserList(UserIndex).Counters.Saliendo = True
    '    UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
    '
    '    Call SendData(ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & IntervaloCerrarConexion & " segundos..." & FONTTYPE_INFO)
    ' [/OLD]
    'ElseIf Not UserList(UserIndex).Counters.Saliendo Then
    '    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    '    Call SendData(ToIndex, UserIndex, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
    '    Call SendData(ToIndex, UserIndex, 0, "FINOK")
    '
    '    Call CloseUser(UserIndex)
    '    UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    '    frmMain.Socket2(UserIndex).Cleanup
    '    Unload frmMain.Socket2(UserIndex)
    '    Call ResetUserSlot(UserIndex)
    End If
End Sub

Public Sub GanarExp(ByVal UserIndex As Integer, ByVal exp As Long, ByVal EsMascota As Boolean)
On Error Resume Next
If EstaEnParty(UserIndex) Then
    If EsLiderParty(UserIndex) = False Then
        Call PartyExp(UserList(UserIndex).flags.LiderParty, exp)
    Else
        Call PartyExp(UserIndex, exp)
    End If
Else
        Call AddtoVar(UserList(UserIndex).Stats.exp, exp, MaxExp)
        If EsMascota = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & exp & " puntos de experiencia." & FONTTYPE_FIGHT_MASCOTA)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & exp & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
        End If
        Call CheckUserLevel(UserIndex)
End If
End Sub

' 0.12b3
Public Function QuitarObj(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
Dim i As Integer
QuitarObj = False
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Call Desequipar(UserIndex, i)
        UserList(UserIndex).Invent.Object(i).Amount = 0
        UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        QuitarObj = True
    End If
Next i
For i = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(UserIndex).BancoInvent.Object(i).ObjIndex = ItemIndex Then
        UserList(UserIndex).BancoInvent.Object(i).ObjIndex = 0
        UserList(UserIndex).BancoInvent.Object(i).Amount = 0
        QuitarObj = True
    End If
Next i
If QuitarObj = True Then Call UpdateUserInv(True, UserIndex, 0)
End Function
