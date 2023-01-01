Attribute VB_Name = "Admin"
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

Public Type tMotd
    Texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloMover As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION) ' Si no es la version 1
If VersionOK = True Then Exit Function
VersionOK = (Ver = ULTIMAVERSION2) ' Tal vez, sea la 2 :P
End Function


Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error GoTo erroraka

Dim i As Integer
Dim MiNPC As Npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(i, 0)
        End If
   End If
   
Next i
Exit Sub
erroraka:
    Call LogError("Error en ReSpawnOrigPosNPCs - Err: " & Err.Number & " Desc:" & Err.Description)

End Sub

Sub WorldSave()
' [GS]
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long


DoEvents
Call SendData(ToAdmins, 0, 0, "||--- ReSwap Guardias..." & "~32~51~223~0~0")
Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
Call SendData(ToAdmins, 0, 0, "||OK" & "~223~51~32~1~1")
Call SendData(ToAdmins, 0, 0, "||--- Localizando cambios en los mapas..." & "~32~51~223~0~0")
Dim j As Integer, k As Integer
For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j
Call SendData(ToAdmins, 0, 0, "||OK" & "~223~51~32~1~1")

' [GS] Algo para entretenerse
For j = 1 To LastUser
    If UserList(j).ConnID <> -1 Then
        Call SendMOTD(j)
    End If
Next
' [/GS]

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.Value = 0

Call SendData(ToAll, 0, 0, "||--- Guardando mapas..." & "~32~51~223~0~0")
For loopX = 1 To NumMaps
    'DoEvents
    
    If MapInfo(loopX).BackUp = 1 Then
        If MapInfo(loopX).Cargado = True Then
            If MAPA_PRETORIANO <> loopX Then
                Call SaveMapData(loopX)
            End If
        End If
        FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
    End If

Next loopX
Call SendData(ToAll, 0, 0, "||OK" & "~223~51~32~1~1")
FrmStat.Visible = False

Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_VENENO)

Call SendData(ToAll, 0, 0, "||--- Guardando NPCs..." & "~32~51~223~0~0")
If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next
Call SendData(ToAll, 0, 0, "||OK" & "~223~51~32~1~1")
Call SendData(ToAll, 0, 0, "||WORLDSAVE DONE" & FONTTYPE_WARNING)
' [/GS]
End Sub

Public Sub PurgarPenas()
On Error Resume Next
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged And UserList(i).ConnID <> -1 Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(ToIndex, i, 0, "||Has sido liberado!" & FONTTYPE_INFO)
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
' [OLD]
'If UCase$(Name) = "GS" Then
'    BANCheck = False
'    Exit Function
'ElseIf UCase$(Name) = "GS" Then
'    BANCheck = False
'    Exit Function
'End If
' [/OLD]
' [NEW]
If Inbaneable(Name) Then Exit Function
' [/NEW]

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1) 'Or _
(val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NOONE")
    UnBan = True
Else
    UnBan = False
End If
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(MD5s)
        If (md5formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = True
End Function
