Attribute VB_Name = "modClanes"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


Option Explicit

Public Guilds As New Collection



Public Sub ComputeVote(ByVal UserIndex As Integer, ByVal rdata As String)

Dim myGuild As cGuild

Set myGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If myGuild Is Nothing Then Exit Sub

If Not myGuild.Elections Then
   Call SendData(ToIndex, UserIndex, 0, "||Aun no es periodo de elecciones." & FONTTYPE_GUILD)
   Exit Sub
End If

If UserList(UserIndex).GuildInfo.YaVoto = 1 Then
   Call SendData(ToIndex, UserIndex, 0, "||Ya has votado!!! solo se permite un voto por miembro." & FONTTYPE_GUILD)
   Exit Sub
End If

If Not myGuild.IsMember(rdata) Then
   Call SendData(ToIndex, UserIndex, 0, "||No hay ningun miembro con ese nombre." & FONTTYPE_GUILD)
   Exit Sub
End If


Call myGuild.Votes.Add(rdata)
UserList(UserIndex).GuildInfo.YaVoto = 1
Call SendData(ToIndex, UserIndex, 0, "||Tu voto ha sido contabilizado." & FONTTYPE_GUILD)


End Sub

Public Sub ResetUserVotes(ByRef myGuild As cGuild)

On Error GoTo errh

Dim k As Integer, Index As Integer
Dim UserFile As String
For k = 1 To myGuild.Members.Count
       
    Index = DameUserIndexConNombre(myGuild.Members(k))
    If Index <> 0 Then 'is online
        UserList(Index).GuildInfo.YaVoto = 0
    Else
        UserFile = CharPath & UCase$(myGuild.Members(k)) & ".chr"
        If FileExist(UserFile, vbNormal) Then
                Call WriteVar(UserFile, "GUILD", "YaVoto", 0)
        End If
    End If
    
Next k

errh:

End Sub


Public Sub DayElapsed()

On Error GoTo errh

Dim T%
Dim MemberIndex As Integer
Dim UserFile As String

For T% = 1 To Guilds.Count
    
    If Guilds(T%).DaysSinceLastElection < Guilds(T%).ElectionPeriod Then
        Guilds(T%).DaysSinceLastElection = Guilds(T%).DaysSinceLastElection + 1
    Else
       If Guilds(T%).Elections = False Then
            Guilds(T%).ResetVotes
            Call ResetUserVotes(Guilds(T%))
            Guilds(T%).Elections = True
            
            MemberIndex = DameGuildMemberIndex(Guilds(T%).GuildName)
            
            If MemberIndex <> 0 Then
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Hoy es la votacion para elegir un nuevo lider para el clan!!." & FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||La eleccion durara 24 horas, se puede votar a cualquier miembro del clan." & FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Para votar escribe /VOTO NICKNAME." & FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Solo se computara un voto por miembro." & FONTTYPE_GUILD)
            End If
        Else
            If Guilds(T%).Members.Count > 1 Then
                    'compute elections results
                    Dim leader$, newleaderindex As Integer, oldleaderindex As Integer
                    leader$ = Guilds(T%).NuevoLider
                    Guilds(T%).Elections = False
                    MemberIndex = DameGuildMemberIndex(Guilds(T%).GuildName)
                    newleaderindex = DameUserIndexConNombre(leader$)
                    oldleaderindex = DameUserIndexConNombre(Guilds(T%).leader)
                    
                    If UCase$(leader$) <> UCase$(Guilds(T%).leader) Then
                        
                        
                        
                        If oldleaderindex <> 0 Then
                            UserList(oldleaderindex).GuildInfo.EsGuildLeader = 0
                        Else
                            UserFile = CharPath & UCase$(Guilds(T%).leader) & ".chr"
                            If FileExist(UserFile, vbNormal) Then
                                    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", 0)
                            End If
                        End If
                        
                        If newleaderindex <> 0 Then
                            UserList(newleaderindex).GuildInfo.EsGuildLeader = 1
                            Call AddtoVar(UserList(newleaderindex).GuildInfo.VecesFueGuildLeader, 1, 10000)
                        Else
                            UserFile = CharPath & UCase$(leader$) & ".chr"
                            If FileExist(UserFile, vbNormal) Then
                                    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", 1)
                            End If
                        End If
                        
                        Guilds(T%).leader = leader$
                    End If
                    
                    If MemberIndex <> 0 Then
                            Call SendData(ToGuildMembers, MemberIndex, 0, "||La elecciones han finalizado!!." & FONTTYPE_GUILD)
                            Call SendData(ToGuildMembers, MemberIndex, 0, "||El nuevo lider es " & leader$ & FONTTYPE_GUILD)
                    End If
                    
                    If newleaderindex <> 0 Then
                        Call SendData(ToIndex, newleaderindex, 0, "||���Has ganado las elecciones, felicitaciones!!!" & FONTTYPE_GUILD)
                        Call GiveGuildPoints(400, newleaderindex)
                    End If
                    Guilds(T%).DaysSinceLastElection = 0
            End If
        End If
    End If
    
Next T%

Exit Sub

errh:
    Call LogError(Err.Description & " error en DayElapsed.")

End Sub

Public Sub GiveGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)

If SendNotice Then _
   Call SendData(ToIndex, UserIndex, 0, "||���Has recibido " & Pts & " guildpoints!!!" & FONTTYPE_GUILD)

Call AddtoVar(UserList(UserIndex).GuildInfo.GuildPoints, Pts, 9000000)

End Sub

Public Sub DropGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)

UserList(UserIndex).GuildInfo.GuildPoints = UserList(UserIndex).GuildInfo.GuildPoints - Pts

'If UserList(UserIndex).GuildInfo.GuildPoints < (-5000) Then
'
'End If

End Sub


Public Sub AcceptPeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(rdata)

If oGuild Is Nothing Then Exit Sub

If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||No estas en guerra con el clan." & FONTTYPE_GUILD)
    Exit Sub
End If

Call oGuild.RemoveEnemy(UserList(UserIndex).GuildInfo.GuildName)

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Call oGuild.RemoveEnemy(rdata)
Call oGuild.RemoveProposition(rdata)

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(rdata)

If MemberIndex <> 0 Then _
    Call SendData(ToGuildMembers, MemberIndex, 0, "||El clan firm� la paz con " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_GUILD)
    
Call SendData(ToGuildMembers, UserIndex, 0, "||El clan firm� la paz con " & rdata & FONTTYPE_GUILD)




End Sub


Public Sub SendPeaceRequest(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud

Set Soli = oGuild.GetPeaceRequest(rdata)

If Soli Is Nothing Then Exit Sub

Call SendData(ToIndex, UserIndex, 0, "PEACEDE" & Soli.desc)

End Sub


Public Sub RecievePeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim h$

h$ = UCase$(ReadField(1, rdata, 44))

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(h$) Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(h$)

If oGuild Is Nothing Then Exit Sub

If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||No estas en guerra con el clan." & FONTTYPE_GUILD)
    Exit Sub
End If

If oGuild.IsAllie(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||Ya estas en paz con el clan." & FONTTYPE_GUILD)
    Exit Sub
End If

Dim peaceoffer As New cSolicitud

peaceoffer.desc = ReadField(2, rdata, 44)
peaceoffer.UserName = UserList(UserIndex).GuildInfo.GuildName

If Not oGuild.IncludesPeaceOffer(peaceoffer.UserName) Then
    Call oGuild.PeacePropositions.Add(peaceoffer)
    Call SendData(ToIndex, UserIndex, 0, "||La propuesta de paz ha sido entregada." & FONTTYPE_GUILD)
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya has enviado una propuesta de paz." & FONTTYPE_GUILD)
End If


End Sub


Public Sub SendPeacePropositions(ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim L%, k$

If oGuild.PeacePropositions.Count = 0 Then Exit Sub

k$ = "PEACEPR" & oGuild.PeacePropositions.Count & ","

For L% = 1 To oGuild.PeacePropositions.Count
    k$ = k$ & oGuild.PeacePropositions(L%).UserName & ","
Next L%

Call SendData(ToIndex, UserIndex, 0, k$)

End Sub

Public Sub AutoEacharMember(ByVal UserIndex As Integer)

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then
    Call SendData(ToIndex, UserIndex, 0, "||No eres miembro de ningun clan." & FONTTYPE_GUILD)
    Exit Sub
End If

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(UserList(UserIndex).Name)
If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
    'Call SendData(ToIndex, UserIndex, 0, "||No te puedes echar a ti mismo si eres el lider del clan." & FONTTYPE_INFO)
    Call BORRARClan(UserIndex)
    Exit Sub
End If

If MemberIndex <> 0 Then 'esta online
    Call SendData(ToIndex, MemberIndex, 0, "||Has sido expulsado del clan." & FONTTYPE_GUILD)
    Call AddtoVar(UserList(MemberIndex).GuildInfo.Echadas, 1, 1000)
    UserList(MemberIndex).GuildInfo.GuildPoints = 0
    UserList(MemberIndex).GuildInfo.GuildName = ""
    ' [GS] Corrige error de mapa
    Call ResetUserChar(ToMap, 0, UserList(MemberIndex).Pos.Map, MemberIndex)
    ' [/GS]
End If

Call oGuild.RemoveMember(UserList(MemberIndex).Name)

End Sub

' [NEW] Hiper-AO
'Public Sub AutoEacharMember(ByVal Userindex As Integer)

'On Error GoTo FallaAutoEchar

'Dim oGuild As cGuild

'Set oGuild = FetchGuild(UserList(Userindex).GuildInfo.GuildName)

'If oGuild Is Nothing Then
'    Call SendData(ToIndex, Userindex, 0, "||No eres miembro de ningun clan." & FONTTYPE_GUILD)
'    Exit Sub
'End If

'Dim MemberIndex As Integer

'MemberIndex = DameUserIndexConNombre(UserList(Userindex).Name)

'If MemberIndex <> 0 Then 'esta online
'    Call AddtoVar(UserList(MemberIndex).GuildInfo.Echadas, 1, 1000)
'    UserList(MemberIndex).GuildInfo.GuildPoints = 0
'    UserList(MemberIndex).GuildInfo.GuildName = ""
'    ' [GS] Para que el Lider se expulse!! Solo si esta solo en el clan!
'    If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
'        UserList(MemberIndex).GuildInfo.EsGuildLeader = 0
'        UserList(MemberIndex).GuildInfo.ClanFundado = ""
'    End If
'    'If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 And oGuild.Members.Count = 1 Then
'    '    UserList(MemberIndex).GuildInfo.EsGuildLeader = 0
'    '    UserList(MemberIndex).GuildInfo.ClanFundado = ""
'    'Else
'    '    Call SendData(ToIndex, MemberIndex, 0, "||No puedes expulsarte, tienes que expulsar a todos los demas integrantes." & FONTTYPE_GUILD)
'    '    Exit Sub
'    'End If
'    ' [/GS]
'    UserList(MemberIndex).GuildInfo.FundoClan = 0
'    Call SendData(ToIndex, MemberIndex, 0, "||Has sido expulsado del clan." & FONTTYPE_GUILD)
'End If

'Call oGuild.RemoveMember(UserList(MemberIndex).Name)
'Exit Sub

'FallaAutoEchar:
'    Call LogError("Error en AutoEcharMiembro del clan - Num: " & Err.Number & " Desc: " & Err.Description)


'End Sub
' [/NEW]



Public Sub EacharMember(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(rdata)
If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No te puedes echar a ti mismo si eres el lider del clan. Si deseas borrar el clan, escribe /salirclan." & FONTTYPE_INFO)
    Exit Sub
End If
If MemberIndex <> 0 Then 'esta online
    Call SendData(ToIndex, MemberIndex, 0, "||Has sido expulsado del clan." & FONTTYPE_GUILD)
    Call AddtoVar(UserList(MemberIndex).GuildInfo.Echadas, 1, 1000)
    UserList(MemberIndex).GuildInfo.GuildPoints = 0
    UserList(MemberIndex).GuildInfo.GuildName = ""
    ' [GS] Corrige error de mapa
    Call ResetUserChar(ToMap, 0, UserList(MemberIndex).Pos.Map, MemberIndex)
    ' [/GS]
    Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue expulsado del clan." & FONTTYPE_GUILD)
Else
    Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta ONLINE." & FONTTYPE_GUILD)
    Exit Sub
End If

Call oGuild.RemoveMember(UserList(MemberIndex).Name)



End Sub

Public Sub BORRARClan(UserIndex As Integer)
On Error GoTo Fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
    If UserList(UserIndex).GuildInfo.BorroClan = True Then
        Call SendData(ToIndex, UserIndex, 0, "||No esta permitido abusar de este comando." & FONTTYPE_GUILD)
        Call SendData(ToIndex, UserIndex, 0, "||Debes esperar unos minutos para volver a utilizar este comando." & FONTTYPE_GUILD)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " intento borrar clan, ya abiendo borrado o creado otro clan hace menor de un minuto." & FONTTYPE_ADMIN)
        Call LogCOSAS("Clanes", UserList(UserIndex).Name & " intento borrar clan, ya abiendo borrado o creado otro clan hace menos de un minuto.")
        UserList(UserIndex).GuildInfo.BorroClan = True
        Exit Sub
    End If
    Dim oGuild As cGuild
    Dim tIndex As Integer
    Dim ClanX As String
    ClanX = UserList(UserIndex).GuildInfo.GuildName
    Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
    If oGuild Is Nothing Then Exit Sub
    Dim j, k As Integer
    For j = 1 To Guilds.Count
        If UCase$(Guilds(j).GuildName) = UCase$(UserList(UserIndex).GuildInfo.GuildName) Then
            For k = 1 To oGuild.Members.Count
                tIndex = NameIndex(UCase$(oGuild.Members(k)))
                If tIndex <> 0 Then
                    'Call SendData(ToIndex, tIndex, 0, "||" & ClanX & " se ha desintegrado." & FONTTYPE_GUILD)
                    Call SendData(ToIndex, tIndex, 0, "||Has sido expulsado del clan." & FONTTYPE_GUILD)
                    Call AddtoVar(UserList(tIndex).GuildInfo.Echadas, 1, 1000)
                    UserList(tIndex).GuildInfo.GuildPoints = 0
                    UserList(tIndex).GuildInfo.GuildName = ""
                    UserList(tIndex).GuildInfo.FundoClan = 0
                    ' [GS] Corrige error de mapa
                    Call ResetUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex)
                    ' [/GS]
                    Call AddtoVar(UserList(tIndex).GuildInfo.ClanesParticipo, 1, 10000)
                    Call oGuild.RemoveMember(UCase$(oGuild.Members(k)))
                End If
            Next k
        UserList(UserIndex).GuildInfo.BorroClan = True
        Call AddtoVar(UserList(UserIndex).GuildInfo.VecesFueGuildLeader, 1, 10000)
        UserList(UserIndex).GuildInfo.ClanFundado = ""
        UserList(UserIndex).GuildInfo.EsGuildLeader = 0
        Call Guilds.Remove(j)
        Dim f As String
        f = App.Path & "\Guilds\" & ClanX & "-Allied" & ".all"
        If FileExist(f, vbNormal) Then Kill f
        f = App.Path & "\Guilds\" & ClanX & "-Enemys" & ".ene"
        If FileExist(f, vbNormal) Then Kill f
        f = App.Path & "\Guilds\" & ClanX & "-Members" & ".mem"
        If FileExist(f, vbNormal) Then Kill f
        f = App.Path & "\Guilds\" & ClanX & "-Solicitudes" & ".sol"
        If FileExist(f, vbNormal) Then Kill f
        f = App.Path & "\Guilds\" & ClanX & "-Propositions" & ".pro"
        If FileExist(f, vbNormal) Then Kill f
        Call SendData(ToAll, 0, 0, "||El clan '" & ClanX & "' a dejado de existir." & FONTTYPE_GUILD)
        Exit Sub
        End If
    Next j
End If
Exit Sub
Fallo:
    Call LogError("Error al Borrar Clan " & ClanX & " - " & Err.Number & " - " & Err.Description)
End Sub



' ### BORRANDO CLANES ###

Public Sub SakarMember(ByVal UserIndex As Integer, ByVal rdata As String)
On Error Resume Next
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If ExisteGuild(UserList(UserIndex).GuildInfo.GuildName) = False Then Exit Sub
If oGuild Is Nothing Then Exit Sub
Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(rdata)
UserList(MemberIndex).GuildInfo.GuildPoints = 0
UserList(MemberIndex).GuildInfo.GuildName = ""
UserList(MemberIndex).GuildInfo.FundoClan = 0
UserList(MemberIndex).GuildInfo.EsGuildLeader = 0
Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue removido del clan." & FONTTYPE_GUILD)
Call oGuild.RemoveMember(UserList(MemberIndex).Name)

End Sub

' ### BORRANDO CLANES ###




Public Sub DenyRequest(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(Soli.UserName)

If MemberIndex <> 0 Then 'esta online
    Call SendData(ToIndex, MemberIndex, 0, "||Tu solicitud ha sido rechazada." & FONTTYPE_GUILD)
    Call AddtoVar(UserList(MemberIndex).GuildInfo.SolicitudesRechazadas, 1, 10000)
End If

Call oGuild.RemoveSolicitud(Soli.UserName)

End Sub


Public Sub AcceptClanMember(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(Soli.UserName)

If MemberIndex <> 0 Then 'esta online
    
    If UserList(MemberIndex).GuildInfo.GuildName <> "" Then
        Call SendData(ToIndex, UserIndex, 0, "||No pod�s aceptar esa solicitud, el pesonaje es lider de otro clan." & FONTTYPE_GUILD)
        Exit Sub
    End If
    
    UserList(MemberIndex).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName
    Call AddtoVar(UserList(MemberIndex).GuildInfo.ClanesParticipo, 1, 1000)
    Call SendData(ToIndex, MemberIndex, 0, "||Felicitaciones, tu solicitud ha sido aceptada." & FONTTYPE_GUILD)
    Call SendData(ToIndex, MemberIndex, 0, "||Ahora sos un miembro activo del clan " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_GUILD)
    Call GiveGuildPoints(25, MemberIndex)
    ' [GS] Corrige error de mapa
    Call ResetUserChar(ToMap, 0, UserList(MemberIndex).Pos.Map, MemberIndex)
    ' [/GS]
Else
    Call SendData(ToIndex, UserIndex, 0, "||Solo podes aceptar solicitudes cuando el solicitante esta ONLINE." & FONTTYPE_GUILD)
    Exit Sub
End If

Call oGuild.Members.Add(Soli.UserName)
Call oGuild.RemoveSolicitud(Soli.UserName)
Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_ACEPTADOCLAN)
Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " ha sido aceptado en el clan." & FONTTYPE_GUILD)


End Sub


Public Sub SendPeticion(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
    
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

  
Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Call SendData(ToIndex, UserIndex, 0, "PETICIO" & Soli.desc)


End Sub


Public Sub SolicitudIngresoClan(ByVal UserIndex As Integer, ByVal Data As String)

If EsNewbie(UserIndex) Then
   Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden conformar clanes." & FONTTYPE_GUILD)
   Exit Sub
End If

Dim MiSol As New cSolicitud

MiSol.desc = ReadField(2, Data, 44)
MiSol.UserName = UserList(UserIndex).Name

Dim clan$

clan$ = ReadField(1, Data, 44)


Dim oGuild As cGuild

Set oGuild = FetchGuild(clan$)

If oGuild Is Nothing Then Exit Sub

If oGuild.IsMember(UserList(UserIndex).Name) Then Exit Sub


If Not oGuild.SolicitudesIncludes(MiSol.UserName) Then
        Call AddtoVar(UserList(UserIndex).GuildInfo.Solicitudes, 1, 1000)
        
        Call oGuild.TestSolicitudBound
        Call oGuild.Solicitudes.Add(MiSol)
        
        Call SendData(ToIndex, UserIndex, 0, "||La solicitud fue recibida por el lider del clan, ahora debes esperar la respuesta." & FONTTYPE_GUILD)
        Exit Sub
Else
        Call SendData(ToIndex, UserIndex, 0, "||Tu solicitud ya fue recibida por el lider del clan, ahora debes esperar la respuesta." & FONTTYPE_GUILD)
End If


End Sub


Public Sub SendCharInfo(ByVal UserName As String, ByVal UserIndex As Integer)

'�Existe el personaje?

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub


Dim UserFile As String
UserFile = CharPath & UCase$(UserName) & ".chr"

If FileExist(UserFile, vbNormal) = False Then Exit Sub

Dim MiUser As User

MiUser.Name = UserName
MiUser.raza = Raza2Num(GetVar(UserFile, "INIT", "Raza"))
MiUser.clase = Clase2Num(GetVar(UserFile, "INIT", "Clase"))
MiUser.genero = Gen2Num(GetVar(UserFile, "INIT", "Genero"))
MiUser.Stats.ELV = val(GetVar(UserFile, "STATS", "ELV"))
MiUser.Stats.GLD = val(GetVar(UserFile, "STATS", "GLD"))
MiUser.Stats.banco = val(GetVar(UserFile, "STATS", "BANCO"))
MiUser.Reputacion.Promedio = val(GetVar(UserFile, "REP", "Promedio"))

Dim h$
h$ = "CHRINFO" & UserName & ","
h$ = h$ & Num2Raza(MiUser.raza) & ","
h$ = h$ & Num2Clase(MiUser.clase) & ","
h$ = h$ & Num2Gen(MiUser.genero) & ","
h$ = h$ & MiUser.Stats.ELV & ","
h$ = h$ & MiUser.Stats.GLD & ","
h$ = h$ & MiUser.Stats.banco & ","
h$ = h$ & MiUser.Reputacion.Promedio & ","


MiUser.GuildInfo.FundoClan = val(GetVar(UserFile, "Guild", "FundoClan"))
MiUser.GuildInfo.EsGuildLeader = val(GetVar(UserFile, "Guild", "EsGuildLeader"))
MiUser.GuildInfo.Echadas = val(GetVar(UserFile, "Guild", "Echadas"))
MiUser.GuildInfo.Solicitudes = val(GetVar(UserFile, "Guild", "Solicitudes"))
MiUser.GuildInfo.SolicitudesRechazadas = val(GetVar(UserFile, "Guild", "SolicitudesRechazadas"))
MiUser.GuildInfo.VecesFueGuildLeader = val(GetVar(UserFile, "Guild", "VecesFueGuildLeader"))
'MiUser.GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
MiUser.GuildInfo.ClanesParticipo = val(GetVar(UserFile, "Guild", "ClanesParticipo"))

h$ = h$ & MiUser.GuildInfo.FundoClan & ","
h$ = h$ & MiUser.GuildInfo.EsGuildLeader & ","
h$ = h$ & MiUser.GuildInfo.Echadas & ","
h$ = h$ & MiUser.GuildInfo.Solicitudes & ","
h$ = h$ & MiUser.GuildInfo.SolicitudesRechazadas & ","
h$ = h$ & MiUser.GuildInfo.VecesFueGuildLeader & ","
h$ = h$ & MiUser.GuildInfo.ClanesParticipo & ","


MiUser.GuildInfo.ClanFundado = GetVar(UserFile, "Guild", "ClanFundado")
MiUser.GuildInfo.GuildName = GetVar(UserFile, "Guild", "GuildName")


h$ = h$ & MiUser.GuildInfo.ClanFundado & ","
h$ = h$ & MiUser.GuildInfo.GuildName & ","


MiUser.Faccion.ArmadaReal = val(GetVar(UserFile, "FACCIONES", "EjercitoReal"))
MiUser.Faccion.FuerzasCaos = val(GetVar(UserFile, "FACCIONES", "EjercitoCaos"))
MiUser.Faccion.CiudadanosMatados = val(GetVar(UserFile, "FACCIONES", "CiudMatados"))
MiUser.Faccion.CriminalesMatados = val(GetVar(UserFile, "FACCIONES", "CrimMatados"))

h$ = h$ & MiUser.Faccion.ArmadaReal & ","
h$ = h$ & MiUser.Faccion.FuerzasCaos & ","
h$ = h$ & MiUser.Faccion.CiudadanosMatados & ","


Call SendData(ToIndex, UserIndex, 0, h$)


End Sub



Public Sub UpdateGuildNews(ByVal rdata As String, ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

oGuild.GuildNews = rdata

End Sub


Public Sub UpdateCodexAndDesc(ByVal rdata As String, ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Call oGuild.UpdateCodexAndDesc(rdata)

End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub


Dim cad$, T%

'<-------Lista de guilds ---------->

cad$ = "LEADERI" & Guilds.Count & "�"

For T% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(T%).GuildName & "�"
Next T%

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub


'<-------Lista de miembros ---------->

cad$ = cad$ & oGuild.Members.Count & "�"

For T% = 1 To oGuild.Members.Count
    cad$ = cad$ & oGuild.Members.Item(T%) & "�"
Next T%


'<------- Guild News -------->

Dim GN$

GN$ = Replace(oGuild.GuildNews, vbCrLf, "�")

cad$ = cad$ & GN$ & "�"

'<------- Solicitudes ------->

cad$ = cad$ & oGuild.Solicitudes.Count & "�"

For T% = 1 To oGuild.Solicitudes.Count
    cad$ = cad$ & oGuild.Solicitudes.Item(T%).UserName & "�"
Next T%

Call SendData(ToIndex, UserIndex, 0, cad$)


End Sub

Public Sub SetNewURL(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

oGuild.URL = rdata

Call SendData(ToIndex, UserIndex, 0, "||La direccion de la web ha sido actualizada" & FONTTYPE_INFO)

End Sub

Public Sub DeclareAllie(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub


Dim LeaderGuild As cGuild, enemyGuild As cGuild

Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub

Set enemyGuild = FetchGuild(rdata)

If enemyGuild Is Nothing Then Exit Sub

If LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        Call SendData(ToIndex, UserIndex, 0, "||Estas en guerra con �ste clan, antes debes firmar la paz." & FONTTYPE_GUILD)
Else
   If Not LeaderGuild.IsAllie(enemyGuild.GuildName) Then
        Call LeaderGuild.AlliedGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.AlliedGuilds.Add(LeaderGuild.GuildName)
        
        Call SendData(ToGuildMembers, UserIndex, 0, "||Tu clan ha firmado una alianza con " & enemyGuild.GuildName & FONTTYPE_GUILD)
        Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        
        Dim Index As Integer
        Index = DameGuildMemberIndex(enemyGuild.GuildName)
        If Index <> 0 Then
            Call SendData(ToGuildMembers, Index, 0, "||" & LeaderGuild.GuildName & " firmo una alianza con tu clan." & FONTTYPE_GUILD)
            Call SendData(ToGuildMembers, Index, 0, "TW" & SND_DECLAREWAR)
        End If
   Else
        Call SendData(ToIndex, UserIndex, 0, "||Ya estas aliado con �ste clan." & FONTTYPE_GUILD)
   End If
End If

    


End Sub


Public Sub DeclareWar(ByVal UserIndex As Integer, ByVal rdata As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub


Dim LeaderGuild As cGuild, enemyGuild As cGuild

Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub

Set enemyGuild = FetchGuild(rdata)

If enemyGuild Is Nothing Then Exit Sub

If Not LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        
        Call LeaderGuild.RemoveAllie(enemyGuild.GuildName)
        Call enemyGuild.RemoveAllie(LeaderGuild.GuildName)
        
        Call LeaderGuild.EnemyGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.EnemyGuilds.Add(LeaderGuild.GuildName)
        
        
        Call SendData(ToGuildMembers, UserIndex, 0, "||Tu clan le declar� la guerra a " & enemyGuild.GuildName & FONTTYPE_GUILD)
        Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        
        Dim Index As Integer
        Index = DameGuildMemberIndex(enemyGuild.GuildName)
        If Index <> 0 Then
            Call SendData(ToGuildMembers, Index, 0, "||" & LeaderGuild.GuildName & " le declarad� la guerra a tu clan." & FONTTYPE_GUILD)
            Call SendData(ToGuildMembers, Index, 0, "TW" & SND_DECLAREWAR)
        End If
Else
   Call SendData(ToIndex, UserIndex, 0, "||Tu clan ya esta en guerra con " & enemyGuild.GuildName & FONTTYPE_GUILD)
End If


End Sub

Public Function DameGuildMemberIndex(ByVal GuildName As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
GuildName = UCase$(GuildName)
  
Do Until UCase$(UserList(LoopC).GuildInfo.GuildName) = GuildName

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameGuildMemberIndex = 0
        Exit Function
    End If
    
Loop
  
DameGuildMemberIndex = LoopC



End Function


Public Sub SendGuildNews(ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub


Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim k$

k$ = "GUILDNE" & oGuild.GuildNews & "�"

Dim T%

k$ = k$ & oGuild.EnemyGuilds.Count & "�"

For T% = 1 To oGuild.EnemyGuilds.Count

    k$ = k$ & oGuild.EnemyGuilds(T%) & "�"
    
Next T%

k$ = k$ & oGuild.AlliedGuilds.Count & "�"

For T% = 1 To oGuild.AlliedGuilds.Count

    k$ = k$ & oGuild.AlliedGuilds(T%) & "�"
    
Next T%



Call SendData(ToIndex, UserIndex, 0, k$)

If oGuild.Elections Then
    Call SendData(ToIndex, UserIndex, 0, "||Hoy es la votacion para elegir un nuevo lider para el clan!!." & FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||La eleccion durara 24 horas, se puede votar a cualquier miembro del clan." & FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||Para votar escribe /VOTO NICKNAME." & FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||Solo se computara un voto por miembro." & FONTTYPE_GUILD)
End If


End Sub

Public Sub SendGuildsList(ByVal UserIndex As Integer)

Dim cad$, T%

cad$ = "GL" & Guilds.Count & ","

For T% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(T%).GuildName & ","
Next T%

Call SendData(ToIndex, UserIndex, 0, cad$)

End Sub

Public Function FetchGuild(ByVal GuildName As String) As Object
Dim k As Integer
For k = 1 To Guilds.Count
    If UCase$(Guilds.Item(k).GuildName) = UCase$(GuildName) Then
            Set FetchGuild = Guilds.Item(k)
            Exit Function
    End If
Next k

Set FetchGuild = Nothing

End Function

Public Sub LoadGuildsDB()

Dim file As String, Cant As Integer

file = App.Path & "\Guilds\" & "GuildsInfo.inf"

If Not FileExist(file, vbNormal) Then Exit Sub

Cant = val(GetVar(file, "INIT", "NroGuilds"))

Dim NewGuild As cGuild
Dim k%

For k% = 1 To Cant
    Set NewGuild = New cGuild
    Call NewGuild.InitializeGuildFromDisk(k%)
    Call Guilds.Add(NewGuild)
Next k%


End Sub

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String)
On Error GoTo errhandler

Dim oGuild As cGuild

If Guilds.Count = 0 Then Exit Sub

Set oGuild = FetchGuild(GuildName)

If oGuild Is Nothing Then Exit Sub

Dim cad$

cad$ = "CLANDET"

cad$ = cad$ & oGuild.GuildName
cad$ = cad$ & "�" & oGuild.Founder
cad$ = cad$ & "�" & oGuild.FundationDate
cad$ = cad$ & "�" & oGuild.leader
cad$ = cad$ & "�" & oGuild.URL
cad$ = cad$ & "�" & oGuild.Members.Count
cad$ = cad$ & "�" & oGuild.DaysToNextElection
cad$ = cad$ & "�" & oGuild.Gold
cad$ = cad$ & "�" & oGuild.EnemyGuilds.Count
cad$ = cad$ & "�" & oGuild.AlliedGuilds.Count

Dim codex$

codex$ = oGuild.CodexLenght()

Dim k%

For k% = 0 To oGuild.CodexLenght()
    codex$ = codex$ & "�" & oGuild.GetCodex(k%)
Next k%


cad$ = cad$ & "�" & codex$ & oGuild.Description


Call SendData(ToIndex, UserIndex, 0, cad$)

errhandler:

End Sub


Public Function CanCreateGuild(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Stats.ELV < NivelMinimoParaFundar Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan debes de ser nivel " & NivelMinimoParaFundar & " o superior" & FONTTYPE_GUILD)
    Exit Function
End If

If UserList(UserIndex).Stats.UserSkills(Liderazgo) < 90 Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan necesitas al menos 90 pts en liderazgo" & FONTTYPE_GUILD)
    Exit Function
End If

' [GS] Para que no salgan por tiempo inactivo tanto
UserList(UserIndex).Counters.IdleCount = 0
' [/GS]

CanCreateGuild = True

End Function

Public Function ExisteGuild(ByVal Name As String) As Boolean

Dim k As Integer
Name = UCase$(Name)

For k = 1 To Guilds.Count
    If UCase$(Guilds(k).GuildName) = Name Then
            ExisteGuild = True
            Exit Function
    End If
Next k

End Function

Public Function CreateGuild(ByVal Name As String, ByVal Rep As Long, ByVal Index As Integer, ByVal GuildInfo As String) As Boolean

' [GS] Para que no salgan por tiempo inactivo tanto
UserList(Index).Counters.IdleCount = 0
' [/GS]

If Not CanCreateGuild(Index) Then
    CreateGuild = False
    Exit Function
End If

Dim miClan As New cGuild

If Not miClan.Initialize(GuildInfo, Name, Rep) Then
    CreateGuild = False
    Call SendData(ToIndex, Index, 0, "||Los datos del clan son invalidos, asegurate que no contiene caracteres invalidos." & FONTTYPE_GUILD)
    Exit Function
End If

If ExisteGuild(miClan.GuildName) Then
    CreateGuild = False
    Call SendData(ToIndex, Index, 0, "||Ya ex�ste un clan con ese nombre." & FONTTYPE_GUILD)
    Exit Function
End If

' [GS] Special
If Right(miClan.GuildName, 1) = " " Then
    Call SendData(ToIndex, Index, 0, "||Nombre invalido, remueva los espacios al final del nombre." & FONTTYPE_GUILD)
    Exit Function
End If

If Left(miClan.GuildName, 1) = " " Then
    Call SendData(ToIndex, Index, 0, "||Nombre invalido, remueva los espacios al principio del nombre." & FONTTYPE_GUILD)
    Exit Function
End If
' [/GS]

Call miClan.Members.Add(UCase$(UserList(Index).Name))

Call Guilds.Add(miClan, miClan.GuildName)

UserList(Index).GuildInfo.FundoClan = 1
UserList(Index).GuildInfo.EsGuildLeader = 1

Call AddtoVar(UserList(Index).GuildInfo.VecesFueGuildLeader, 1, 10000)
Call AddtoVar(UserList(Index).GuildInfo.ClanesParticipo, 1, 10000)

UserList(Index).GuildInfo.ClanFundado = miClan.GuildName
UserList(Index).GuildInfo.GuildName = UserList(Index).GuildInfo.ClanFundado

Call GiveGuildPoints(5000, Index)

Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||���" & UserList(Index).Name & " fundo el clan '" & UserList(Index).GuildInfo.GuildName & "'!!!" & FONTTYPE_GUILD)

UserList(Index).GuildInfo.BorroClan = True

CreateGuild = True

End Function


Public Sub SaveGuildsDB()
On Error GoTo Fallo

Dim j As Integer
Dim file As String

file = App.Path & "\Guilds\" & "GuildsInfo.inf"

If FileExist(file, vbNormal) Then Kill file

Call WriteVar(file, "INIT", "NroGuilds", str(Guilds.Count))

For j = 1 To Guilds.Count
    
    Call Guilds(j).SaveGuild(file, j)
    
Next j

' 0.12b3
If Guilds.Count > 1 Then
    file = App.Path & "\Guilds\" & "GuildsInfo.bak"
    If FileExist(file, vbNormal) Then Kill file
    Call WriteVar(file, "INIT", "NroGuilds", str(Guilds.Count))
    For j = 1 To Guilds.Count
        Call Guilds(j).SaveGuild(file, j)
    Next j
End If

Exit Sub

Fallo:
    Call LogError("Error al Guardar Clanes - " & Err.Number & " - " & Err.Description)

End Sub
