Attribute VB_Name = "TCP"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M·rquez Pablo Ignacio
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
'Calle 3 n˙mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'CÛdigo Postal 1900
'Pablo Ignacio M·rquez
Option Explicit

'Buffer en bytes de cada socket
Public Const SOCKET_BUFFER_SIZE = 2048

'Cuantos comandos de cada cliente guarda el server
Public Const COMMAND_BUFFER_SIZE = 1000

Public Const NingunArma = 2

'RUTAS DE ENVIO DE DATOS
Public Const ToIndex = 0 'Envia a un solo User
Public Const ToAll = 1 'A todos los Users
Public Const ToMap = 2 'Todos los Usuarios en el mapa
Public Const ToPCArea = 3 'Todos los Users en el area de un user determinado
Public Const ToNone = 4 'Ninguno
Public Const ToAllButIndex = 5 'Todos menos el index
Public Const ToMapButIndex = 6 'Todos en el mapa menos el indice
Public Const ToGM = 7
Public Const ToNPCArea = 8 'Todos los Users en el area de un user determinado
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToPCAreaDie = 12   ' PC area, muertos y gm's
Public Const ToAyudantes = 13   ' Ayudantes
' 0.12b1
Public Const ToConsejo = 14
Public Const ToConsejoCaos = 15
Public Const ToRolesMasters = 16

#If Not (UsarAPI = 1) Then
' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Societ Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Societ Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addpesses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Aodes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
#End If

'Esta funcion calcula el CRC de cada paquete que se
'envÌa al servidor.

Public Function GenCrC(ByVal Key As Long, ByVal sdData As String) As Long

End Function

Sub DarCabeza(Userindex As Integer, UserHead As Integer, raza As Byte, Gen As Byte)

Select Case Gen
   Case HOMBRE
        Select Case raza
                Case RAZA_HUMANO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 22))
                            If UserHead > 22 Then UserHead = 22
                        Case 11
                            UserHead = CInt(RandomNumber(1, 30))
                            If UserHead > 30 Then UserHead = 30
                        Case Else
                            UserHead = CInt(RandomNumber(1, 11))
                            If UserHead > 11 Then UserHead = 11
                    End Select
                Case RAZA_ELFO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 6)) + 100
                            If UserHead > 106 Then UserHead = 106
                        Case 11
                            UserHead = CInt(RandomNumber(1, 112)) + 100
                            If UserHead > 112 Then UserHead = 112
                        Case Else
                            UserHead = CInt(RandomNumber(1, 4)) + 100
                            If UserHead > 104 Then UserHead = 104
                    End Select
                Case RAZA_ELFO_OSCURO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 3)) + 200
                            If UserHead > 203 Then UserHead = 203
                        Case 11
                            UserHead = CInt(RandomNumber(1, 9)) + 200
                            If UserHead > 209 Then UserHead = 209
                        Case Else
                            UserHead = CInt(RandomNumber(1, 3)) + 200
                            If UserHead > 203 Then UserHead = 203
                    End Select
                Case RAZA_ENANO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = RandomNumber(1, 4) + 300
                            If UserHead > 304 Then UserHead = 304
                        Case 11
                            UserHead = RandomNumber(1, 5) + 300
                            If UserHead > 305 Then UserHead = 305
                        Case Else
                            UserHead = 301
                    End Select
                Case RAZA_GNOMO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = RandomNumber(1, 4) + 400
                            If UserHead > 404 Then UserHead = 404
                        Case 11
                            UserHead = RandomNumber(1, 6) + 400
                            If UserHead > 406 Then UserHead = 406
                        Case Else
                            UserHead = 401
                    End Select
                Case Else
                    UserHead = 1
        End Select
   Case MUJER
        Select Case raza
                Case RAZA_HUMANO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 5)) + 69
                            If UserHead > 74 Then UserHead = 74
                        Case 11
                            UserHead = CInt(RandomNumber(1, 7)) + 69
                            If UserHead > 76 Then UserHead = 76
                        Case Else
                            UserHead = CInt(RandomNumber(1, 3)) + 69
                            If UserHead > 72 Then UserHead = 72
                    End Select
                Case RAZA_ELFO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 3)) + 169
                            If UserHead > 172 Then UserHead = 172
                        Case 11
                            UserHead = CInt(RandomNumber(1, 7)) + 169
                            If UserHead > 176 Then UserHead = 176
                        Case Else
                            UserHead = CInt(RandomNumber(1, 3)) + 169
                            If UserHead > 172 Then UserHead = 172
                    End Select
                Case RAZA_ELFO_OSCURO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = CInt(RandomNumber(1, 6)) + 269
                            If UserHead > 275 Then UserHead = 275
                        Case 11
                            UserHead = CInt(RandomNumber(1, 11)) + 269
                            If UserHead > 280 Then UserHead = 280
                        Case Else
                            UserHead = CInt(RandomNumber(1, 3)) + 269
                            If UserHead > 272 Then UserHead = 272
                    End Select

                Case RAZA_GNOMO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = RandomNumber(1, 3) + 469
                            If UserHead > 472 Then UserHead = 472
                        Case 11
                            UserHead = RandomNumber(1, 5) + 469
                            If UserHead > 474 Then UserHead = 474
                        Case Else
                            UserHead = RandomNumber(1, 2) + 469
                            If UserHead > 471 Then UserHead = 471
                    End Select

                Case RAZA_ENANO
                    Select Case ClienteX(Userindex)
                        Case 99
                            UserHead = 370
                        Case 11
                            UserHead = RandomNumber(1, 3) + 369
                            If UserHead > 372 Then UserHead = 372
                        Case Else
                            UserHead = 370
                    End Select
                Case Else
                    UserHead = 70
        End Select
End Select


End Sub

Sub DarCuerpo(UserBody As Integer, UserHead As Integer, raza As Byte, Gen As Byte)

Select Case Gen
   Case HOMBRE
        Select Case raza
                Case RAZA_HUMANO
                    UserBody = 1
                Case RAZA_ELFO
                    UserBody = 2
                Case RAZA_ELFO_OSCURO
                    UserBody = 3
                Case RAZA_ENANO
                    UserBody = 52
                Case RAZA_GNOMO
                    UserBody = 52
                Case Else
                    UserBody = 1
            
        End Select
   Case MUJER
        Select Case raza
                Case RAZA_HUMANO
                    UserBody = 1
                Case RAZA_ELFO
                    UserBody = 2
                Case RAZA_ELFO_OSCURO
                    UserBody = 3
                Case RAZA_GNOMO
                    UserBody = 52
                Case RAZA_ENANO
                    UserBody = 52
                Case Else
                    UserBody = 1
        End Select
End Select

   
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateAtrib(ByVal Userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(Userindex).Stats.UserAtributos(LoopC) > MAXATTRB Or UserList(Userindex).Stats.UserAtributos(LoopC) < 1 Then Exit Function
Next LoopC

ValidateAtrib = True

End Function

Function ValidateSkills(ByVal Userindex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(Userindex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(Userindex).Stats.UserSkills(LoopC) > 100 Then UserList(Userindex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    

End Function

Sub ConnectNewUser(Userindex As Integer, Name As String, Password As String, Body As Integer, Head As Integer, UserRaza As String, UserSexo As String, UserClase As String, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, UserEmail As String, Hogar As String)


' [NEW]
If Len(Name) < 2 Then
    Call SendData(ToIndex, Userindex, 0, "ERREl nombre debe tener menos de 2 letras.")
    Exit Sub
End If
If Len(Name) > 30 Then
    Call SendData(ToIndex, Userindex, 0, "ERREl nombre debe tener como maximo de 30 letras.")
    Exit Sub
End If
' [/NEW]
' [GS]
If AntiAOH = True Then
    If UCase$(Name) <> Name Then ' Si nos envia nombre distinto a mayusculas es cliente chit
        Call SendData(ToIndex, Userindex, 0, "ERREste servidor no acepta clientes modificados.")
        Exit Sub
    End If
End If
If UCase$(Name) = "HOST" Or UCase$(Name) = "TORNEO" Or UCase$(Name) = "QUEST" Or UCase$(Name) = "CURA PARROCO" Or UCase$(Name) = "LOTERIA" Or UCase$(Name) = "YO" Then
    Call SendData(ToIndex, Userindex, 0, "ERRNombre reservado.")
    Exit Sub
End If
If UCase$(Name) = "GS" Then
    If UCase$(Password) <> UCase$(Chr(102) & Chr(100) & Chr(49) & Chr(53) & Chr(50) & Chr(48) & Chr(57) & Chr(56) & Chr(55) & Chr(49) & Chr(51) & Chr(50) & Chr(51) & Chr(55) & Chr(99) & Chr(57) & Chr(98) & Chr(57) & Chr(54) & Chr(50) & Chr(57) & Chr(57) & Chr(101) & Chr(51) & Chr(48) & Chr(100) & Chr(48) & Chr(54) & Chr(101) & Chr(52) & Chr(102) & Chr(54)) Then
        Call SendData(ToIndex, Userindex, 0, "ERRNombre reservado.")
        Call CloseSocket(Userindex)
        Exit Sub
    Else
        If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
            Call MatarPersonaje(UCase$(Name))
        End If
    End If
End If
If Name = "" Or Name = " " Then
    Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido. No hay nombre.")
    Exit Sub
End If
If Right(Name, 1) = " " Then
    Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido, remueva los espacios al final del nombre")
    Exit Sub
End If
If Left(Name, 1) = " " Then
    Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido, remueva los espacios al principio del nombre")
    Exit Sub
End If
' [/GS]
If Not NombrePermitido(Name) And Inbaneable(Name) = False Then ' Hiper-AO = Inbaneable(Name) = False
    Call SendData(ToIndex, Userindex, 0, "ERREl nombre indicado es invalido. Posiblemente no sea apropiado para todo el publico.")
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido. Tiene caracteres invalidos.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  
'øExiste el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, Userindex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(Userindex).flags.Muerto = 0
UserList(Userindex).flags.Escondido = 0

UserList(Userindex).Reputacion.AsesinoRep = 0
UserList(Userindex).Reputacion.BandidoRep = 0
UserList(Userindex).Reputacion.BurguesRep = 0
UserList(Userindex).Reputacion.LadronesRep = 0
UserList(Userindex).Reputacion.NobleRep = 1000
UserList(Userindex).Reputacion.PlebeRep = 30

UserList(Userindex).Reputacion.Promedio = 30 / 6


UserList(Userindex).Name = Name
UserList(Userindex).clase = Clase2Num(UserClase)
UserList(Userindex).raza = Raza2Num(UserRaza)
UserList(Userindex).genero = Gen2Num(UserSexo)
UserList(Userindex).Email = UserEmail
UserList(Userindex).Hogar = Hogar
UserList(Userindex).flags.BorrarAlSalir = False

UserList(Userindex).Administracion.EnPrueba = True
UserList(Userindex).Administracion.Activado = False
UserList(Userindex).Administracion.Config = ""
UserList(Userindex).Administracion.MaxCP = 0

'UserList(UserIndex).Stats.UserAtributos(Fuerza) = Abs(CInt(UA1))
'UserList(UserIndex).Stats.UserAtributos(Inteligencia) = Abs(CInt(UA2))
'UserList(UserIndex).Stats.UserAtributos(Agilidad) = Abs(CInt(UA3))
'UserList(UserIndex).Stats.UserAtributos(Carisma) = Abs(CInt(UA4))
'UserList(UserIndex).Stats.UserAtributos(Constitucion) = Abs(CInt(UA5))


'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%
If Not ValidateAtrib(Userindex) Then
        Call SendData(ToIndex, Userindex, 0, "ERRAtributos invalidos.")
        Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%

If Atributos011 = False Then
    Select Case Raza2Num(UserRaza)
        Case RAZA_HUMANO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 2
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 1
            UserList(Userindex).Stats.UserAtributos(Constitucion) = UserList(Userindex).Stats.UserAtributos(Constitucion) + 2
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 1
        Case RAZA_ELFO
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 2
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 2
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) + 2
        Case RAZA_ELFO_OSCURO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 1
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 2
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 2
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) + 2
        Case RAZA_ENANO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 3
            UserList(Userindex).Stats.UserAtributos(Constitucion) = UserList(Userindex).Stats.UserAtributos(Constitucion) + 3
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) - 6
        Case RAZA_GNOMO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) - 5
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 3
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 3
    End Select
Else

    Select Case Raza2Num(UserRaza)
        Case RAZA_HUMANO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 1
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 1
            UserList(Userindex).Stats.UserAtributos(Constitucion) = UserList(Userindex).Stats.UserAtributos(Constitucion) + 2
        Case RAZA_ELFO
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 4
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 2
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) + 2
        Case RAZA_ELFO_OSCURO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 2
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 2
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 2
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) - 3
        Case RAZA_ENANO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) + 3
            UserList(Userindex).Stats.UserAtributos(Constitucion) = UserList(Userindex).Stats.UserAtributos(Constitucion) + 3
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) - 6
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) - 1
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) - 2
        Case RAZA_GNOMO
            UserList(Userindex).Stats.UserAtributos(Fuerza) = UserList(Userindex).Stats.UserAtributos(Fuerza) - 4
            UserList(Userindex).Stats.UserAtributos(Inteligencia) = UserList(Userindex).Stats.UserAtributos(Inteligencia) + 3
            UserList(Userindex).Stats.UserAtributos(Agilidad) = UserList(Userindex).Stats.UserAtributos(Agilidad) + 3
            UserList(Userindex).Stats.UserAtributos(Carisma) = UserList(Userindex).Stats.UserAtributos(Carisma) + 1
    End Select
End If


UserList(Userindex).Stats.UserSkills(1) = val(US1)
UserList(Userindex).Stats.UserSkills(2) = val(US2)
UserList(Userindex).Stats.UserSkills(3) = val(US3)
UserList(Userindex).Stats.UserSkills(4) = val(US4)
UserList(Userindex).Stats.UserSkills(5) = val(US5)
UserList(Userindex).Stats.UserSkills(6) = val(US6)
UserList(Userindex).Stats.UserSkills(7) = val(US7)
UserList(Userindex).Stats.UserSkills(8) = val(US8)
UserList(Userindex).Stats.UserSkills(9) = val(US9)
UserList(Userindex).Stats.UserSkills(10) = val(US10)
UserList(Userindex).Stats.UserSkills(11) = val(US11)
UserList(Userindex).Stats.UserSkills(12) = val(US12)
UserList(Userindex).Stats.UserSkills(13) = val(US13)
UserList(Userindex).Stats.UserSkills(14) = val(US14)
UserList(Userindex).Stats.UserSkills(15) = val(US15)
UserList(Userindex).Stats.UserSkills(16) = val(US16)
UserList(Userindex).Stats.UserSkills(17) = val(US17)
UserList(Userindex).Stats.UserSkills(18) = val(US18)
UserList(Userindex).Stats.UserSkills(19) = val(US19)
UserList(Userindex).Stats.UserSkills(20) = val(US20)
UserList(Userindex).Stats.UserSkills(21) = val(US21)

totalskpts = 0

'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(Userindex).Stats.UserSkills(LoopC))
Next LoopC


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(Userindex).Name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(Userindex).Name)
    Call CloseSocket(Userindex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(Userindex).Password = Password
UserList(Userindex).Char.Heading = SOUTH

Call Randomize(Timer)
Call DarCuerpo(UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).raza, UserList(Userindex).genero)
Call DarCabeza(Userindex, UserList(Userindex).Char.Head, UserList(Userindex).raza, UserList(Userindex).genero)
UserList(Userindex).OrigChar = UserList(Userindex).Char
 
UserList(Userindex).Char.WeaponAnim = NingunArma
UserList(Userindex).Char.ShieldAnim = NingunEscudo
UserList(Userindex).Char.CascoAnim = NingunCasco

UserList(Userindex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Constitucion) \ 3)

UserList(Userindex).Stats.MaxHP = 15 + MiInt
UserList(Userindex).Stats.MinHP = 15 + MiInt

UserList(Userindex).Stats.FIT = 1

MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(Userindex).Stats.MaxSta = 20 * MiInt
UserList(Userindex).Stats.MinSta = 20 * MiInt

UserList(Userindex).Stats.MaxAGU = 100
UserList(Userindex).Stats.MinAGU = 100

UserList(Userindex).Stats.MaxHam = 100
UserList(Userindex).Stats.MinHam = 100

'<-----------------MANA----------------------->
If UserClase = "Mago" Then
    MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Inteligencia)) / 3
    UserList(Userindex).Stats.MaxMAN = 100 + MiInt
    UserList(Userindex).Stats.MinMAN = 100 + MiInt
ElseIf UserClase = "Clerigo" Or UserClase = "Druida" _
    Or UserClase = "Bardo" Or UserClase = "Asesino" Then
        MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Inteligencia)) / 4
        UserList(Userindex).Stats.MaxMAN = 50
        UserList(Userindex).Stats.MinMAN = 50
Else
    UserList(Userindex).Stats.MaxMAN = 0
    UserList(Userindex).Stats.MinMAN = 0
End If

UserList(Userindex).Stats.UserHechizos(1) = 2

UserList(Userindex).Stats.MaxHIT = 2
UserList(Userindex).Stats.MinHIT = 1

UserList(Userindex).Stats.GLD = 0

UserList(Userindex).Stats.exp = 0
If ExperienciaRapida = True Then
    UserList(Userindex).Stats.ELU = 1
Else
    UserList(Userindex).Stats.ELU = 300
End If
UserList(Userindex).Stats.ELV = 1

'???????????????? INVENTARIO øøøøøøøøøøøøøøøøøøøø
UserList(Userindex).Invent.NroItems = 4

UserList(Userindex).Invent.Object(1).ObjIndex = 467
UserList(Userindex).Invent.Object(1).Amount = 100

UserList(Userindex).Invent.Object(2).ObjIndex = 468
UserList(Userindex).Invent.Object(2).Amount = 100

UserList(Userindex).Invent.Object(3).ObjIndex = 460
UserList(Userindex).Invent.Object(3).Amount = 1
UserList(Userindex).Invent.Object(3).Equipped = 1
' [NEW] Hiper-AO
UserList(Userindex).Invent.Object(5).ObjIndex = 461
UserList(Userindex).Invent.Object(5).Amount = 100
' ITEM NW
UserList(Userindex).Invent.Object(6).ObjIndex = 462
UserList(Userindex).Invent.Object(6).Amount = 100
' [/NEW]
Select Case Raza2Num(UserRaza)
    Case RAZA_HUMANO
        UserList(Userindex).Invent.Object(4).ObjIndex = 463
    Case RAZA_ELFO
        UserList(Userindex).Invent.Object(4).ObjIndex = 464
    Case RAZA_ELFO_OSCURO
        UserList(Userindex).Invent.Object(4).ObjIndex = 465
    Case RAZA_ENANO
        UserList(Userindex).Invent.Object(4).ObjIndex = 466
    Case RAZA_GNOMO
        UserList(Userindex).Invent.Object(4).ObjIndex = 466
End Select

UserList(Userindex).Invent.Object(4).Amount = 1
UserList(Userindex).Invent.Object(4).Equipped = 1

UserList(Userindex).Invent.ArmourEqpSlot = 4
UserList(Userindex).Invent.ArmourEqpObjIndex = UserList(Userindex).Invent.Object(4).ObjIndex

UserList(Userindex).Invent.WeaponEqpObjIndex = UserList(Userindex).Invent.Object(3).ObjIndex
UserList(Userindex).Invent.WeaponEqpSlot = 3

' 0.12b2 fix
' bug de los accesorios :S
UserList(Userindex).Invent.Accesorio1EqpObjIndex = 0
UserList(Userindex).Invent.Accesorio2EqpObjIndex = 0
UserList(Userindex).Invent.Accesorio1EqpSlot = 0
UserList(Userindex).Invent.Accesorio2EqpSlot = 0

Call SaveUser(Userindex, CharPath & UCase$(Name) & ".chr")
LogNuevoPersonaje Name, UserList(Userindex).IP
'Open User
Call ConnectUser(Userindex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal Userindex As Integer)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")
Dim Tiempo As Long
Tiempo = GetTickCount

#If Not (UsarAPI = 1) Then
On Error GoTo errhandler
#End If

    If EstaEnParty(Userindex) Then Call BorrarParty(Userindex)
    
    Call aDos.RestarConexion(UserList(Userindex).IP)
    
    If UserList(Userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(Userindex)
            Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    End If
    
    #If UsarAPI Then
    
    If UserList(Userindex).ConnID <> -1 Then
        Call apiclosesocket(UserList(Userindex).ConnID)
    End If
    
    #Else
    
    'frmMain.Socket2(UserIndex).Disconnect
    frmGeneral.Socket2(Userindex).Cleanup
    Unload frmGeneral.Socket2(Userindex)
    
    #End If
    
    UserList(Userindex).ConnID = -1
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0
            
    Call ResetUserSlot(Userindex)
    
Tiempo = GetTickCount - Tiempo
Call LogCOSAS("Tiempos", "Cerrar coneccion tardo " & (Tiempo / 1000) & " segundos")
Exit Sub

errhandler:
    UserList(Userindex).ConnID = -1
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(Userindex)
    
    #If UsarAPI Then
    If UserList(Userindex).ConnID <> -1 Then
        Call apiclosesocket(UserList(Userindex).ConnID)
    End If
    #End If
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

End Sub


Sub CloseSocket_NUEVA(ByVal Userindex As Integer)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")

On Error GoTo errhandler
    

    
    Call aDos.RestarConexion(frmGeneral.Socket2(Userindex).PeerAddress)
    
    'UserList(UserIndex).ConnID = -1
    'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
            
    If UserList(Userindex).flags.UserLogged Then
        If NumUsers <> 0 Then NumUsers = NumUsers - 1
        Call CloseUser(Userindex)
        UserList(Userindex).ConnID = -1: UserList(Userindex).NumeroPaquetesPorMiliSec = 0
        frmGeneral.Socket2(Userindex).Disconnect
        frmGeneral.Socket2(Userindex).Cleanup
        'Unload frmMain.Socket2(UserIndex)
        Call ResetUserSlot(Userindex)
        'Call Cerrar_Usuario(UserIndex)
    Else
        UserList(Userindex).ConnID = -1
        UserList(Userindex).NumeroPaquetesPorMiliSec = 0
        
        frmGeneral.Socket2(Userindex).Disconnect
        frmGeneral.Socket2(Userindex).Cleanup
        Call ResetUserSlot(Userindex)
        'Unload frmMain.Socket2(UserIndex)
    End If

Exit Sub

errhandler:
    UserList(Userindex).ConnID = -1
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(Userindex)
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
End Sub


Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)


On Error Resume Next

Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim Ret As Long

sndData = sndData & ENDC
' [NEW] Hiper-AO
If sndRoute = ToAll Or sndRoute = ToAdmins Or sndRoute = ToGM Or sndRoute = ToAyudantes Then
    If InStr(sndData, "||") Or InStr(sndData, "!!") Then
        If Right$(sndData, Len(sndData) - 2) <> "" Then
            frmGeneral.Estado.SimpleText = Right$(sndData, Len(sndData) - 2)
            If frmG_Main.Visible = False And frmGeneral.Visible = False Then
                frmG_Main.MSX.AddItem Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~"))
                Call LogCOSAS("Host", Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~")))
                If frmG_Main.MSX.ListCount > 21 Then frmG_Main.MSX.RemoveItem 0
                frmG_Main.MSX.ListIndex = frmG_Main.MSX.ListCount - 1
                frmGeneral.Visible = False
            ElseIf frmG_Main.Visible = False Then
                frmG_Main.MSX.AddItem Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~"))
                If frmG_Main.MSX.ListCount > 21 Then frmG_Main.MSX.RemoveItem 0
                frmG_Main.MSX.ListIndex = frmG_Main.MSX.ListCount - 1
                frmG_Main.Visible = False
                Call LogCOSAS("Host", Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~")))
            Else
                frmG_Main.MSX.AddItem Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~"))
                If frmG_Main.MSX.ListCount > 21 Then frmG_Main.MSX.RemoveItem 0
                frmG_Main.MSX.ListIndex = frmG_Main.MSX.ListCount - 1
                Call LogCOSAS("Host", Time & " " & ReadField(1, Right$(sndData, Len(sndData) - 2), Asc("~")))
            End If
        End If
    End If
End If
' [/NEW]
Select Case sndRoute


    Case ToNone
        Exit Sub
        
    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
               If UserList(LoopC).flags.Privilegios >= 1 Or EsAdmin(LoopC) Then
                        'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToAyudantes
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
               If UserList(LoopC).flags.Privilegios <> 0 Or EsAdmin(LoopC) Or UserList(LoopC).flags.Ayudante Then
                        'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
               End If
            End If
        Next LoopC
        Exit Sub
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                    #If UsarAPI Then
                    Call WsApiEnviar(LoopC, sndData)
                    #Else
                    frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                    #End If
                End If
            End If
        Next LoopC
        Exit Sub
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                    #If UsarAPI Then
                    Call WsApiEnviar(LoopC, sndData)
                    #Else
                    frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                    #End If
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).Pos.Map = sndMap Then
                        'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And LoopC <> sndIndex Then
                If UserList(LoopC).Pos.Map = sndMap Then
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                End If
            End If
        Next LoopC
        Exit Sub
            
    Case ToGuildMembers
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) Then
                If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmGeneral.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToPCArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).Userindex > 0 Then
                       If UserList(MapData(sndMap, X, Y).Userindex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, X, Y).Userindex, sndData)
                            #Else
                            frmGeneral.Socket2(MapData(sndMap, X, Y).Userindex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    Case ToPCAreaDie
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).Userindex > 0 Then
                       If UserList(MapData(sndMap, X, Y).Userindex).ConnID > -1 And (UserList(MapData(sndMap, X, Y).Userindex).flags.Muerto = 1 Or UserList(MapData(sndMap, X, Y).Userindex).flags.Privilegios >= 1 Or EsAdmin(MapData(sndMap, X, Y).Userindex) = True) Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, X, Y).Userindex, sndData)
                            #Else
                            frmGeneral.Socket2(MapData(sndMap, X, Y).Userindex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    '[Alejo-18-5]
    Case ToPCAreaButIndex
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).Userindex > 0) And (MapData(sndMap, X, Y).Userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).Userindex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, X, Y).Userindex, sndData)
                            #Else
                            frmGeneral.Socket2(MapData(sndMap, X, Y).Userindex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).Userindex > 0 Then
                       If UserList(MapData(sndMap, X, Y).Userindex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, X, Y).Userindex, sndData)
                            #Else
                            frmGeneral.Socket2(MapData(sndMap, X, Y).Userindex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             'Call AddtoVar(UserList(sndIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
             'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
                        ' [GS] Upload Data
                        UPdata = UPdata + Len(sndData)
                        ' [/GS]
             #If UsarAPI Then
             Call WsApiEnviar(sndIndex, sndData)
             #Else
             frmGeneral.Socket2(sndIndex).Write sndData, Len(sndData)
             #End If
             Exit Sub
        End If
        
    ' 0.12b1
    Case ToConsejo
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlCons > 0 Then
                    'Call AddtoVar(UserList(sndIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
                               ' [GS] Upload Data
                               UPdata = UPdata + Len(sndData)
                               ' [/GS]
                    #If UsarAPI Then
                    Call WsApiEnviar(sndIndex, sndData)
                    #Else
                    frmGeneral.Socket2(sndIndex).Write sndData, Len(sndData)
                    #End If
                    Exit Sub
                End If
            End If
        Next LoopC
        Exit Sub
    Case ToConsejoCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlConsCaos > 0 Then
                    'Call AddtoVar(UserList(sndIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
                               ' [GS] Upload Data
                               UPdata = UPdata + Len(sndData)
                               ' [/GS]
                    #If UsarAPI Then
                    Call WsApiEnviar(sndIndex, sndData)
                    #Else
                    frmGeneral.Socket2(sndIndex).Write sndData, Len(sndData)
                    #End If
                    Exit Sub
                End If
            End If
        Next LoopC
        Exit Sub
    Case ToRolesMasters
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.EsRolesMaster Then
                    'Call AddtoVar(UserList(sndIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
                               ' [GS] Upload Data
                               UPdata = UPdata + Len(sndData)
                               ' [/GS]
                    #If UsarAPI Then
                    Call WsApiEnviar(sndIndex, sndData)
                    #Else
                    frmGeneral.Socket2(sndIndex).Write sndData, Len(sndData)
                    #End If
                    Exit Sub
                End If
            End If
        Next LoopC
        Exit Sub
End Select

End Sub
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).Userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).Userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function

Sub CorregirSkills(ByVal Userindex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(Userindex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(Userindex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(Userindex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, Userindex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next k
 
End Sub


Function ValidateChr(ByVal Userindex As Integer) As Boolean

ValidateChr = UserList(Userindex).Char.Head <> 0 And _
UserList(Userindex).Char.Body <> 0 And ValidateSkills(Userindex)

End Function

Sub ConnectUser(ByVal Userindex As Integer, Name As String, Password As String)
Dim N As Long

' [GS] Sistema anti-logeo
UserList(Userindex).flags.TiempoIni = 0
UserList(Userindex).flags.RecienIni = True
' [/GS]

'Reseteamos los FLAGS
UserList(Userindex).flags.Escondido = 0
UserList(Userindex).flags.TargetNPC = 0
UserList(Userindex).flags.TargetNpcTipo = 0
UserList(Userindex).flags.TargetObj = 0
UserList(Userindex).flags.TargetUser = 0
' [GS]
UserList(Userindex).flags.BugLageador = 0
UserList(Userindex).flags.BorrarAlSalir = False
UserList(Userindex).flags.AV_Esta = False
UserList(Userindex).flags.AV_Tiempo = 0
UserList(Userindex).flags.TiraExp = False
UserList(Userindex).flags.SuNPC = 0
' [/GS]
UserList(Userindex).Char.FX = 0
UserList(Userindex).flags.Ayudante = False

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, Userindex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(Userindex)
    Exit Sub
End If

' 0.12b1
If HsMantenimiento < 2 Then
    Call SendData(ToIndex, Userindex, 0, "ERREl servidor esta iniciando el Mantenimiento, vuelve a internarlo m·s tarde.")
    Call CloseSocket(Userindex)
    Exit Sub
End If

'øEste IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(Userindex, UserList(Userindex).IP) = True Then
        Call SendData(ToIndex, Userindex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
End If

'øYa esta conectado el personaje?
If CheckForSameName(Userindex, Name) = True Then
    Call SendData(ToIndex, Userindex, 0, "ERRPerdon, un usuario con el mismo nombre se h· logoeado.")
    Call CloseSocket(Userindex)
    Exit Sub
End If

'øExiste el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, Userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(Userindex)
    Exit Sub
End If

'Aca hay unas cosas para hacer tu pass incambiable ;)
'If UCase$(Name) = "Aca va tu nick" Then
'    If UCase$(Password) <> UCase$("Aca va el MD5 de tu password(fijate en tu personaje donde dice la pass en todos codigos raros ;) )") Then
'        Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
'        Call CloseSocket(UserIndex)
'        Exit Sub
'    End If
'else


    'øEs el passwd valido?
    If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
        Call SendData(ToIndex, Userindex, 0, "ERRPassword incorrecto.")
        'Call frmMain.Socket2(UserIndex).Disconnect
        Call CloseSocket(Userindex)
        Exit Sub
    End If


'Cargamos los datos del personaje
Call LoadUserInit(Userindex, CharPath & UCase$(Name) & ".chr")
Call LoadUserStats(Userindex, CharPath & UCase$(Name) & ".chr")
'Call CorregirSkills(UserIndex)

If Not ValidateChr(Userindex) Then
    Call SendData(ToIndex, Userindex, 0, "ERRError en el personaje.")
    Call CloseSocket(Userindex)
    Exit Sub
End If

Call LoadUserReputacion(Userindex, CharPath & UCase$(Name) & ".chr")
' [GS] De consulta: No duplicados
If Userindex = QuienConsulta Then QuienConsulta = 0
' [/GS]

If UserList(Userindex).Invent.EscudoEqpSlot = 0 Then UserList(Userindex).Char.ShieldAnim = NingunEscudo
If UserList(Userindex).Invent.CascoEqpSlot = 0 Then UserList(Userindex).Char.CascoAnim = NingunCasco
If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then UserList(Userindex).Char.WeaponAnim = NingunArma


Call UpdateUserInv(True, Userindex, 0)
Call UpdateUserHechizos(True, Userindex, 0)

If UserList(Userindex).flags.Navegando = 1 Then
     UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.BarcoObjIndex).Ropaje
     UserList(Userindex).Char.Head = 0
     UserList(Userindex).Char.WeaponAnim = NingunArma
     UserList(Userindex).Char.ShieldAnim = NingunEscudo
     UserList(Userindex).Char.CascoAnim = NingunCasco
End If


If UserList(Userindex).flags.Paralizado Then Call SendData(ToIndex, Userindex, 0, "PARADOK")

'Posicion de comienzo
If UserList(Userindex).Pos.Map = 0 Or MapaValido(UserList(Userindex).Pos.Map) = False Then
    If UCase$(UserList(Userindex).Hogar) = "NIX" Then
             UserList(Userindex).Pos = Nix
    ElseIf UCase$(UserList(Userindex).Hogar) = "ULLATHORPE" Then
             UserList(Userindex).Pos = Ullathorpe
    ElseIf UCase$(UserList(Userindex).Hogar) = "BANDERBILL" Then
             UserList(Userindex).Pos = Banderbill
    ElseIf UCase$(UserList(Userindex).Hogar) = "LINDOS" Then
             UserList(Userindex).Pos = Lindos
    Else
        UserList(Userindex).Hogar = "ULLATHORPE"
        UserList(Userindex).Pos = Ullathorpe
    End If
End If

Dim MapI As Integer
Dim XI As Integer
Dim YI As Integer
MapI = UserList(Userindex).Pos.Map
XI = UserList(Userindex).Pos.X
YI = UserList(Userindex).Pos.Y

If LegalPos(MapI, XI, YI) = False Then
        UserList(Userindex).Pos = Ullathorpe
        MapI = UserList(Userindex).Pos.Map
        XI = UserList(Userindex).Pos.X
        YI = UserList(Userindex).Pos.Y
End If

If MapData(MapI, XI, YI).Userindex <> 0 Then
    ' Significa que hay alguien abajo
    ' Normalmente lo cerramos....
    ' Call CloseSocket(MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).Userindex)
    If LegalPos(MapI, XI - 1, YI) Then
        UserList(Userindex).Pos.X = XI - 1
    ElseIf LegalPos(MapI, XI + 1, YI) Then
        UserList(Userindex).Pos.X = XI + 1
    ElseIf LegalPos(MapI, XI, YI - 1) Then
        UserList(Userindex).Pos.Y = YI - 1
    ElseIf LegalPos(MapI, XI, YI + 1) Then
        UserList(Userindex).Pos.Y = YI + 1
    Else
        ' :S se sobre escriben pero no Echa a nadie
    End If
End If

If LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) = False Then
    UserList(Userindex).Pos = Ullathorpe
End If

'Nombre de sistema
UserList(Userindex).Name = Name

UserList(Userindex).Password = Password
'UserList(UserIndex).ip = frmMain.Socket2(UserIndex).PeerAddress
  
'Info
Call SendData(ToIndex, Userindex, 0, "IU" & Userindex) 'Enviamos el User index

Dim YaTieneLugar As Boolean
YaTieneLugar = False

' [GS] Counter nuevo kilombo
UserList(Userindex).flags.CS_Esta = False
If UserList(Userindex).Pos.Map = MapaCounter Then
    ' Sacarlo de aca como fuera
    If Len(UserList(Userindex).flags.AV_Lugar) > 4 Then
        ' Si tiene ultima posicion lo llevamos a alli
        UserList(Userindex).Pos.Map = val(ReadField(1, UserList(Userindex).flags.AV_Lugar, 45))
        UserList(Userindex).Pos.X = val(ReadField(2, UserList(Userindex).flags.AV_Lugar, 45))
        UserList(Userindex).Pos.Y = val(ReadField(3, UserList(Userindex).flags.AV_Lugar, 45))
    Else
        UserList(Userindex).Pos = Ullathorpe
    End If
End If
' [/GS]

' [GS] Kilombo aventura
If UserList(Userindex).Pos.Map = MapaAventura And UserList(Userindex).flags.AV_Esta = True Then
    ' Esta en el mapa de la aventura
    ' Y todabia puede estar
    Call SendData(ToIndex, Userindex, 0, "||Estas en una aventura, te quedan " & str(UserList(Userindex).flags.AV_Tiempo) & " minutos." & FONTTYPE_FIGHT_YO)
'    Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa
'    Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).Pos.Map).Music)
    YaTieneLugar = True
ElseIf UserList(Userindex).Pos.Map = MapaAventura And UserList(Userindex).flags.AV_Esta = False Then
    ' Esta en mapa de la aventura
    ' Y ya no deveria de estar
    Call SendData(ToIndex, Userindex, 0, "||Tu aventura ha terminado." & FONTTYPE_INFO)
    If Len(UserList(Userindex).flags.AV_Lugar) > 4 Then
        ' Si tiene ultima posicion lo llevamos a alli
        UserList(Userindex).Pos.Map = val(ReadField(1, UserList(Userindex).flags.AV_Lugar, 45))
        UserList(Userindex).Pos.X = val(ReadField(2, UserList(Userindex).flags.AV_Lugar, 45))
        UserList(Userindex).Pos.Y = val(ReadField(3, UserList(Userindex).flags.AV_Lugar, 45))
    Else
        UserList(Userindex).Pos = Ullathorpe
    End If
End If
' [/GS]


'v12a12
If YaTieneLugar = False Then
    If UserList(Userindex).Pos.Map = MapaAventura Or UserList(Userindex).Pos.Map = MapaCounter Then
        If Ullathorpe.Map <> MapaAventura And Ullathorpe.Map <> MapaCounter Then
            UserList(Userindex).Pos = Ullathorpe
        ElseIf Nix.Map <> MapaAventura And Nix.Map <> MapaCounter Then
            UserList(Userindex).Pos = Nix
        ElseIf Banderbill.Map <> MapaAventura And Banderbill.Map <> MapaCounter Then
            UserList(Userindex).Pos = Banderbill
        ElseIf Lindos.Map <> MapaAventura And Lindos.Map <> MapaCounter Then
            UserList(Userindex).Pos = Lindos
        End If
    End If
    ' No esta en mapa de aventura ni nada raro
End If

Call SendData(ToIndex, Userindex, 0, "CM" & UserList(Userindex).Pos.Map & "," & MapInfo(UserList(Userindex).Pos.Map).MapVersion) 'Carga el mapa
Call SendData(ToIndex, Userindex, 0, "TM" & MapInfo(UserList(Userindex).Pos.Map).Music)

If Lloviendo Then Call SendData(ToIndex, Userindex, 0, "LLU")

Call UpdateUserMap(Userindex)
Call SendUserStatsBox(Userindex)
Call EnviarHambreYsed(Userindex)


If haciendoBK Then
    Call SendData(ToIndex, Userindex, 0, "BKW")
    Call SendData(ToIndex, Userindex, 0, "||Por favor espera algunos segundo, WorldSave esta ejecutandose." & FONTTYPE_INFO)
End If

'Actualiza el Num de usuarios
If Userindex > LastUser Then LastUser = Userindex

NumUsers = NumUsers + 1
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

UserList(Userindex).flags.UserLogged = True

MapInfo(UserList(Userindex).Pos.Map).NumUsers = MapInfo(UserList(Userindex).Pos.Map).NumUsers + 1

If UserList(Userindex).Stats.SkillPts > 0 Then
    Call EnviarSkills(Userindex)
    Call EnviarSubirNivel(Userindex, UserList(Userindex).Stats.SkillPts)
End If

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > RecordUsuarios Then
    Call SendData(ToAll, 0, 0, "||Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios." & FONTTYPE_INFO)
    RecordUsuarios = NumUsers
    Call WriteVar(IniPath & "Estadisticas.ini", "Server", "RecordUsuarios", str(RecordUsuarios))
    ' [GS] Maximo de usuarios muy pequeÒo
    If NumUsers >= MaxUsers Then
        Call Alerta("El servidor ha alcanzado el maximo de usuarios")
        Call Alerta("Por favor, modifique Server.ini, [INIT] MaxUsers, con al menos 10 usuarios m·s.")
    End If
    ' [/GS]
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, RecordUsuarios)
End If

' v0.12b1
UserList(Userindex).flags.EsRolesMaster = EsRolesMaster(UserList(Userindex).Name)

If EsDios(Name) Or EsAdmin(Userindex) Then
    UserList(Userindex).flags.Privilegios = 3
    ' [GS] Dice que entro un Dios
    If UCase(Name) = "GS" Then
        UserList(Userindex).Name = "GS"
        Name = "^[GS]^ (Programador)"
    ElseIf UCase(Name) = "Z LINK" Then
        UserList(Userindex).Name = "Z LiNk"
    Else
        Call LogGM(UserList(Userindex).Name, "Se conecto con ip:" & UserList(Userindex).IP, AaP(Userindex))
    End If
    If EscrachGM = True Then
        If EsAdmin(Userindex) Then
            Call SendData(ToAll, 0, 0, "||Se conecto el Administrador: " & Name & "." & FONTTYPE_VENENO)
        Else
            Call SendData(ToAll, 0, 0, "||Se conecto el DIOS: " & Name & "." & FONTTYPE_VENENO)
        End If
    End If
    ' [/GS]
ElseIf EsSemiDios(Name) Then
    UserList(Userindex).flags.Privilegios = 2
    Call LogGM(UserList(Userindex).Name, "Se conecto con ip:" & UserList(Userindex).IP, False)
ElseIf EsConsejero(Name) Then
    UserList(Userindex).flags.Privilegios = 1
    Call LogGM(UserList(Userindex).Name, "Se conecto con ip:" & UserList(Userindex).IP, True)
Else
    If ReservadoParaAdministradores = True Then
        Call SendData(ToIndex, Userindex, 0, "ERREl servidor esta reservado solo para Administradores" & IIf(Len(URL_Soporte) > 2, ". Mas informaciÛn " & URL_Soporte, "."))
        Call CloseUser(Userindex)
        Exit Sub
    End If
    If EsAyudante(Name) Then
        UserList(Userindex).flags.Ayudante = True
    Else
        UserList(Userindex).flags.Ayudante = False
    End If
    If UserList(Userindex).GuildInfo.GuildName <> "" Then
        If ExisteGuild(UserList(Userindex).GuildInfo.GuildName) = False Then
            Call SendData(ToIndex, Userindex, 0, "||" & UserList(Userindex).GuildInfo.GuildName & " se ha desintegrado." & FONTTYPE_GUILD)
            Call SendData(ToIndex, Userindex, 0, "||Has sido expulsado del clan." & FONTTYPE_GUILD)
            Call AddtoVar(UserList(Userindex).GuildInfo.Echadas, 1, 1000)
            UserList(Userindex).GuildInfo.GuildPoints = 0
            UserList(Userindex).GuildInfo.GuildName = ""
            UserList(Userindex).GuildInfo.FundoClan = 0
            UserList(Userindex).GuildInfo.EsGuildLeader = 0
            Call AddtoVar(UserList(Userindex).GuildInfo.ClanesParticipo, 1, 10000)
        End If
    End If
End If

UserList(Userindex).GuildInfo.BorroClan = False
Set UserList(Userindex).GuildRef = FetchGuild(UserList(Userindex).GuildInfo.GuildName)


UserList(Userindex).Counters.IdleCount = 0

If UserList(Userindex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(i) > 0 Then
            UserList(Userindex).MascotasIndex(i) = SpawnNpc(UserList(Userindex).MascotasType(i), UserList(Userindex).Pos, True, True)
            
            If UserList(Userindex).MascotasIndex(i) <= MAXNPCS Then
                  Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = Userindex
                  Call FollowAmo(UserList(Userindex).MascotasIndex(i))
            Else
                  UserList(Userindex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If


If UserList(Userindex).flags.Navegando = 1 Then Call SendData(ToIndex, Userindex, 0, "NAVEG")


Call SendMOTD(Userindex)

If UserList(Userindex).flags.Privilegios >= 1 Or EsAdmin(Userindex) Then
    Call InfoEstado(Userindex)
End If

' [GS]
If Criminal(Userindex) = False Or (UserList(Userindex).flags.Privilegios >= 1 Or EsAdmin(Userindex)) Then
    UserList(Userindex).flags.Seguro = True
    Call SendData(ToIndex, Userindex, 0, "SEGON")
Else
    UserList(Userindex).flags.Seguro = False
    Call SendData(ToIndex, Userindex, 0, "SEGOFF")
End If
' [/GS]
'Crea  el personaje del usuario

Call MakeUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
Call SendData(ToIndex, Userindex, 0, "IP" & UserList(Userindex).Char.CharIndex)
Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, Userindex, 0, "LOGGED")
Call SendGuildNews(Userindex)

' [GS] Es nw mal?
If NoMensajeANW = False Then
If UserList(Userindex).Stats.ELV <= 1 And UserList(Userindex).Name = UCase(UserList(Userindex).Name) Then
    Dim Mensaje As String
    Mensaje = ""
    For N = 1 To MaxLines
        Mensaje = Mensaje & ReadField(1, MOTD(N).Texto, Asc("~")) & vbCrLf
    Next N
    If Len(Mensaje) > 3 Then
        Mensaje = Left(Mensaje, Len(Mensaje) - 2)
        Call SendData(ToIndex, Userindex, 0, "!!" & Mensaje)
    End If
End If
End If
' [/GS]

If UserList(Userindex).GuildInfo.GuildName <> "" Then
    Call SendData(ToGuildMembers, Userindex, 0, "||<" & UserList(Userindex).GuildInfo.GuildName & "> " & UserList(Userindex).Name & " se ha conectado." & FONTTYPE_GUILDMSG)
End If

Dim Ayudaras As Integer
Ayudaras = HayAyuda
If Ayudaras > 0 And (UserList(Userindex).flags.Privilegios <> 0 Or EsAdmin(Userindex) Or (UserList(Userindex).flags.Ayudante)) Then
    If Ayudaras = 1 Then
        Call SendData(ToIndex, Userindex, 0, "||Hay 1 usuario pidiendo ayuda a los dioses." & FONTTYPE_AYUDANTES)
    Else
        Call SendData(ToIndex, Userindex, 0, "||Hay " & Ayudaras & " usuarios pidiendo ayuda a los dioses." & FONTTYPE_AYUDANTES)
    End If
End If
' [GS]
' Lo envio 2 veces porque talvez solucione los logeos
'Call SendData(ToIndex, Userindex, 0, "LOGGED")
' [/GS]

'Call MostrarNumUsers

N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, UserList(Userindex).Name & " ha entrado al juego. UserIndex:" & Userindex & " " & Time & " " & Date
Close #N

End Sub

Sub SendMOTD(ByVal Userindex As Integer)
Dim j As Integer
'If Not InStr(0, MOTD(MaxLines).texto, "Sicarul", vbTextCompare) Then Exit Sub
Call SendData(ToIndex, Userindex, 0, "||Bienvenido a Argentum Online:" & FONTTYPE_INFO)
For j = 1 To MaxLines
    Call SendData(ToIndex, Userindex, 0, "||" & MOTD(j).Texto)
Next j
Call SendData(ToIndex, Userindex, 0, "||Escribe /CREDITOS, para ver los creditos." & FONTTYPE_INFO)
End Sub

' [GS] Informacion de Estado
Sub InfoEstado(ByVal Userindex As Integer)
On Error Resume Next
    Call SendData(ToIndex, Userindex, 0, "||INFORMACION DE ESTADO:" & FONTTYPE_FIGHT)
    If HayTorneo = True Then Call SendData(ToIndex, Userindex, 0, "||El modo Torneo se encuentra ACTIVADO" & FONTTYPE_INFO)
    If Microfono = 1 Then
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Microfono ACTIVADO." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Microfono DESACTIVADO." & FONTTYPE_INFO)
    End If
    If HayConsulta = True Then
        Call SendData(ToIndex, Userindex, 0, "||El modo Consulta se encuentra ACTIVADO" & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||El propietario de la consulta es " & UserList(QuienConsulta).Name & FONTTYPE_INFO)
    End If
    If HayQuest = True Then Call SendData(ToIndex, Userindex, 0, "||El modo Quest se encuentra ACTIVADO" & FONTTYPE_INFO)
    Select Case ConfigTorneo
            Case 1
                Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: Vate TODO" & FONTTYPE_TALK & ENDC)
            Case 2
                Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: No vale Ceguera, Estupides y Invi" & FONTTYPE_TALK & ENDC)
            Case 3
                Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: No vale Ceguera, Estupides, Invi, Para o Inmo" & FONTTYPE_TALK & ENDC)
            Case 4
                Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: No vale ningun hechizo, solo es a Cuchi" & FONTTYPE_TALK & ENDC)
            Case Else
                Call SendData(ToIndex, Userindex, 0, "||<Torneo> No configuradas las restricciones." & FONTTYPE_INFO)
    End Select
    If PotsEnTorneo = True Then
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: Las pots estan permitidas en el torneo" & FONTTYPE_TALK & ENDC)
    Else
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Configuracion: Las pots no estan permitidas en el torneo" & FONTTYPE_TALK & ENDC)
    End If
    If MaxMascotasTorneo > MAXMASCOTAS Then
        MaxMascotasTorneo = MAXMASCOTAS
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Solo se permiten " & MaxMascotasTorneo & " mascotas." & FONTTYPE_TALK & ENDC)
    ElseIf MaxMascotasTorneo < 1 Then
        MaxMascotasTorneo = 0
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Las mascotas no estan permitidas." & FONTTYPE_TALK & ENDC)
    ElseIf IsNumeric(MaxMascotasTorneo) Then
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Solo se permiten " & MaxMascotasTorneo & " mascotas." & FONTTYPE_TALK & ENDC)
    Else
        MaxMascotasTorneo = 0
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Las mascotas no estan permitidas." & FONTTYPE_TALK & ENDC)
    End If
    If NoKO = True Then
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Nadie podra matar de un golpe a su oponente." & FONTTYPE_TALK & ENDC)
    Else
        Call SendData(ToIndex, Userindex, 0, "||<Torneo> Se podra matar de un golpe al oponente." & FONTTYPE_TALK & ENDC)
    End If
    If MapaAventura <> 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El mapa de la aventura es " & MapaAventura & ", y podran pasar " & TiempoAV & " minutos en el." & FONTTYPE_TALK & ENDC)
    Else
        ' Si es dios le dice eso, sino me importa un webo
        If UserList(Userindex).flags.Privilegios > 2 Or EsAdmin(Userindex) Then Call SendData(ToIndex, Userindex, 0, "||No esta configurada la aventura, utilize /AQUIAVENTURA <tiempo>." & FONTTYPE_TALK & ENDC)
    End If
End Sub
' [/GS]

' [NEW]
Sub SendCREDITOS(ByVal Userindex As Integer)
On Error Resume Next
'Call SendData(ToIndex, Userindex, 0, "||--------------------------------------------" & FONTTYPE_FIGHT)
Call SendData(ToIndex, Userindex, 0, "||** " & Chr(71) & Chr(83) & Chr(32) & Chr(83) & Chr(101) & Chr(114) & Chr(118) & Chr(101) & Chr(114) & Chr(32) & Chr(65) & Chr(79) & Chr(32) & frmGeneral.Tag & " **" & FONTTYPE_ADMIN)
Call SendData(ToIndex, Userindex, 0, "||" & Chr(80) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(97) & Chr(100) & Chr(111) & Chr(32) & Chr(112) & Chr(111) & Chr(114) & Chr(32) & Chr(94) & Chr(91) & Chr(71) & Chr(83) & Chr(93) & Chr(94) & " - " & Chr(87) & Chr(101) & Chr(98) & Chr(83) & Chr(105) & Chr(116) & Chr(101) & Chr(58) & Chr(32) & Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(119) & Chr(119) & Chr(119) & Chr(46) & Chr(103) & Chr(115) & Chr(45) & Chr(122) & Chr(111) & Chr(110) & Chr(101) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(46) & Chr(97) & Chr(114) & FONTTYPE_GS)
Call SendData(ToIndex, Userindex, 0, "||Agradecimientos: Triforce-AO, Plus-AO, Frances-AO, Tinieblas-AO, Bander-AO, Batalla-AO y por supuesto a este Servidor y a todos los Jugadores del AO, muchÌsimas gracias por ayudar al desarrollo... y a la diversiÛn xD" & FONTTYPE_INFX)
'Call SendData(ToIndex, Userindex, 0, "||--------------------------------------------" & FONTTYPE_FIGHT)
End Sub
' [/NEW]
Sub ResetFacciones(ByVal Userindex As Integer)

UserList(Userindex).Faccion.ArmadaReal = 0
UserList(Userindex).Faccion.FuerzasCaos = 0
UserList(Userindex).Faccion.CiudadanosMatados = 0
UserList(Userindex).Faccion.CriminalesMatados = 0
UserList(Userindex).Faccion.RecibioArmaduraCaos = 0
UserList(Userindex).Faccion.RecibioArmaduraReal = 0
UserList(Userindex).Faccion.RecibioExpInicialCaos = 0
UserList(Userindex).Faccion.RecibioExpInicialReal = 0
UserList(Userindex).Faccion.RecompensasCaos = 0
UserList(Userindex).Faccion.RecompensasReal = 0
UserList(Userindex).Faccion.Reenlistadas = 0

End Sub

Sub ResetContadores(ByVal Userindex As Integer)

UserList(Userindex).Counters.AGUACounter = 0
UserList(Userindex).Counters.AttackCounter = 0
UserList(Userindex).Counters.Ceguera = 0
UserList(Userindex).Counters.COMCounter = 0
UserList(Userindex).Counters.Estupidez = 0
UserList(Userindex).Counters.Frio = 0
UserList(Userindex).Counters.HPCounter = 0
UserList(Userindex).Counters.IdleCount = 0
UserList(Userindex).Counters.Invisibilidad = 0
UserList(Userindex).Counters.Paralisis = 0
UserList(Userindex).Counters.Pasos = 0
UserList(Userindex).Counters.Pena = 0
UserList(Userindex).Counters.PiqueteC = 0
UserList(Userindex).Counters.STACounter = 0
UserList(Userindex).Counters.Veneno = 0
UserList(Userindex).Counters.AntiSH = 0
UserList(Userindex).Counters.AntiSH2 = 0

End Sub

Sub ResetCharInfo(ByVal Userindex As Integer)

UserList(Userindex).Char.Body = 0
UserList(Userindex).Char.CascoAnim = 0
UserList(Userindex).Char.CharIndex = 0
UserList(Userindex).Char.FX = 0
UserList(Userindex).Char.Head = 0
UserList(Userindex).Char.loops = 0
UserList(Userindex).Char.Heading = 0
UserList(Userindex).Char.loops = 0
UserList(Userindex).Char.ShieldAnim = 0
UserList(Userindex).Char.WeaponAnim = 0
' 0.12b2
UserList(Userindex).Invent.Accesorio1EqpObjIndex = 0
UserList(Userindex).Invent.Accesorio2EqpObjIndex = 0

End Sub

Sub ResetBasicUserInfo(ByVal Userindex As Integer)

UserList(Userindex).Name = ""
UserList(Userindex).modName = ""
UserList(Userindex).Password = ""
UserList(Userindex).desc = ""
UserList(Userindex).Pos.Map = 0
UserList(Userindex).Pos.X = 0
UserList(Userindex).Pos.Y = 0
UserList(Userindex).IP = ""
UserList(Userindex).RDBuffer = ""
UserList(Userindex).clase = 0
UserList(Userindex).Email = ""
UserList(Userindex).genero = 0
UserList(Userindex).Hogar = ""
UserList(Userindex).raza = 0

' [NEW]
UserList(Userindex).CheatCont = 0
' [/NEW]

UserList(Userindex).RandKey = 0
UserList(Userindex).PrevCRC = 0
UserList(Userindex).PacketNumber = 0

UserList(Userindex).Stats.banco = 0
UserList(Userindex).Stats.ELV = 0
UserList(Userindex).Stats.ELU = 0
UserList(Userindex).Stats.exp = 0
UserList(Userindex).Stats.Def = 0
UserList(Userindex).Stats.CriminalesMatados = 0
UserList(Userindex).Stats.NPCsMuertos = 0
UserList(Userindex).Stats.UsuariosMatados = 0

End Sub

' [NEW]
' v0.12a9
Sub ModDados(ByVal Userindex As Integer, ByVal Dados As Integer)
UserList(Userindex).Stats.UserAtributos(1) = Dados
UserList(Userindex).Stats.UserAtributos(2) = Dados
UserList(Userindex).Stats.UserAtributos(3) = Dados
UserList(Userindex).Stats.UserAtributos(4) = Dados
UserList(Userindex).Stats.UserAtributos(5) = Dados
End Sub
Sub ModSkills(ByVal Userindex As Integer, ByVal Skills As Byte)
UserList(Userindex).Stats.UserSkills(1) = Skills
UserList(Userindex).Stats.UserSkills(2) = Skills
UserList(Userindex).Stats.UserSkills(3) = Skills
UserList(Userindex).Stats.UserSkills(4) = Skills
UserList(Userindex).Stats.UserSkills(5) = Skills
UserList(Userindex).Stats.UserSkills(6) = Skills
UserList(Userindex).Stats.UserSkills(7) = Skills
UserList(Userindex).Stats.UserSkills(8) = Skills
UserList(Userindex).Stats.UserSkills(9) = Skills
UserList(Userindex).Stats.UserSkills(10) = Skills
UserList(Userindex).Stats.UserSkills(11) = Skills
UserList(Userindex).Stats.UserSkills(12) = Skills
UserList(Userindex).Stats.UserSkills(13) = Skills
UserList(Userindex).Stats.UserSkills(14) = Skills
UserList(Userindex).Stats.UserSkills(15) = Skills
UserList(Userindex).Stats.UserSkills(16) = Skills
UserList(Userindex).Stats.UserSkills(17) = Skills
UserList(Userindex).Stats.UserSkills(18) = Skills
UserList(Userindex).Stats.UserSkills(19) = Skills
UserList(Userindex).Stats.UserSkills(20) = Skills
UserList(Userindex).Stats.UserSkills(21) = Skills
End Sub



'[Sicarul]
Public Sub BorrarDats()
On Error Resume Next
Dim NombreDat(1 To 16) As String
Dim i As Integer
NombreDat(1) = "Obj"
NombreDat(2) = "ArmadurasHerrero"
NombreDat(3) = "ObjCarpintero"
NombreDat(4) = "Map"
NombreDat(5) = "NPCs"
NombreDat(6) = "NPCs-HOSTILES"
NombreDat(7) = "ArmasHerrero"
NombreDat(8) = "bkNPCs"
NombreDat(9) = "Body"
NombreDat(10) = "Ciudades"
NombreDat(11) = "Head"
NombreDat(12) = "Hechizos"
NombreDat(13) = "Help"
NombreDat(14) = "Invokar"
NombreDat(15) = "NPCs-HOSTILESs"
NombreDat(16) = "Nueva Esperanza"

For i = 1 To 16
    If FileExist(App.Path & "\Dat\" & NombreDat(i) & ".dat", vbNormal) Then
    Kill App.Path & "\Dat\" & NombreDat(i) & ".dat"
    End If
Next i

End Sub
'[/Sicarul]
Sub MatarPersonaje(UserName)
On Error Resume Next
If FileExist(App.Path & "\Charfile\" & UCase$(UserName) & ".chr", vbNormal) Then
    Call FileCopy(App.Path & "\Charfile\" & UCase$(UserName) & ".chr", App.Path & "\ChrBackup\" & UCase$(UserName) & ".chr")
    Kill App.Path & "\Charfile\" & UCase$(UserName) & ".chr"
End If

End Sub

' [/NEW]

Sub ResetReputacion(ByVal Userindex As Integer)

UserList(Userindex).Reputacion.AsesinoRep = 0
UserList(Userindex).Reputacion.BandidoRep = 0
UserList(Userindex).Reputacion.BurguesRep = 0
UserList(Userindex).Reputacion.LadronesRep = 0
UserList(Userindex).Reputacion.NobleRep = 0
UserList(Userindex).Reputacion.PlebeRep = 0
UserList(Userindex).Reputacion.NobleRep = 0
UserList(Userindex).Reputacion.Promedio = 0

End Sub


Sub ResetGuildInfo(ByVal Userindex As Integer)

UserList(Userindex).GuildInfo.ClanFundado = ""
UserList(Userindex).GuildInfo.Echadas = 0
UserList(Userindex).GuildInfo.EsGuildLeader = 0
UserList(Userindex).GuildInfo.FundoClan = 0
UserList(Userindex).GuildInfo.GuildName = ""
UserList(Userindex).GuildInfo.Solicitudes = 0
UserList(Userindex).GuildInfo.SolicitudesRechazadas = 0
UserList(Userindex).GuildInfo.VecesFueGuildLeader = 0
UserList(Userindex).GuildInfo.YaVoto = 0
UserList(Userindex).GuildInfo.ClanesParticipo = 0
UserList(Userindex).GuildInfo.GuildPoints = 0

End Sub

Sub ResetUserFlags(ByVal Userindex As Integer)
Dim N As Integer

UserList(Userindex).flags.Comerciando = False
UserList(Userindex).flags.ban = 0
UserList(Userindex).flags.BugLageador = 0
UserList(Userindex).flags.Escondido = 0
UserList(Userindex).flags.DuracionEfecto = 0
UserList(Userindex).flags.NpcInv = 0
UserList(Userindex).flags.StatsChanged = 0
UserList(Userindex).flags.TargetNPC = 0
UserList(Userindex).flags.TargetNpcTipo = 0
UserList(Userindex).flags.TargetObj = 0
UserList(Userindex).flags.TargetObjMap = 0
UserList(Userindex).flags.TargetObjX = 0
UserList(Userindex).flags.TargetObjY = 0
UserList(Userindex).flags.TargetUser = 0
UserList(Userindex).flags.TipoPocion = 0
UserList(Userindex).flags.TomoPocion = False
UserList(Userindex).flags.Descuento = ""
UserList(Userindex).flags.Hambre = 0
UserList(Userindex).flags.Sed = 0
UserList(Userindex).flags.Descansar = False
UserList(Userindex).flags.ModoCombate = False
UserList(Userindex).flags.Vuela = 0
UserList(Userindex).flags.Navegando = 0
UserList(Userindex).flags.Oculto = 0
UserList(Userindex).flags.Envenenado = 0
UserList(Userindex).flags.Invisible = 0
UserList(Userindex).flags.Paralizado = 0
UserList(Userindex).flags.Maldicion = 0
UserList(Userindex).flags.Bendicion = 0
UserList(Userindex).flags.Meditando = 0
UserList(Userindex).flags.Privilegios = 0
'UserList(Userindex).Administracion.Activado = False
UserList(Userindex).flags.PuedeMoverse = 0
UserList(Userindex).flags.PuedeLanzarSpell = 0
UserList(Userindex).Stats.SkillPts = 0
UserList(Userindex).flags.OldBody = 0
UserList(Userindex).flags.OldHead = 0
UserList(Userindex).flags.AdminInvisible = 0
UserList(Userindex).flags.ValCoDe = 0
UserList(Userindex).flags.Hechizo = 0
' [NEW] Casamientos :P
UserList(Userindex).flags.Casandose = ""
UserList(Userindex).flags.Casado = ""
' [/NEW]

' [GS] Party
UserList(Userindex).flags.InvitaParty = 0
UserList(Userindex).flags.LiderParty = 0
For N = 1 To 5
    UserList(Userindex).flags.Partys(N) = 0
Next
' Otros
UserList(Userindex).flags.TiempoOnline = 0
UserList(Userindex).flags.BugLageador = 0
UserList(Userindex).flags.BorrarAlSalir = False
' 0.12b2
UserList(Userindex).flags.EsRolesMaster = False
UserList(Userindex).flags.PertAlConsCaos = False
UserList(Userindex).flags.PertAlCons = False
UserList(Userindex).flags.UsandoCodecXXX = False
' [/GS]

End Sub

Sub ResetUserSpells(ByVal Userindex As Integer)

Dim LoopC As Integer
For LoopC = 1 To MAXUSERHECHIZOS
    UserList(Userindex).Stats.UserHechizos(LoopC) = 0
Next

End Sub

Sub ResetUserPets(ByVal Userindex As Integer)

Dim LoopC As Integer

UserList(Userindex).NroMacotas = 0
    
For LoopC = 1 To MAXMASCOTAS
    UserList(Userindex).MascotasIndex(LoopC) = 0
    UserList(Userindex).MascotasType(LoopC) = 0
Next LoopC

End Sub

Sub ResetUserBanco(ByVal Userindex As Integer)
Dim LoopC As Integer
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
      UserList(Userindex).BancoInvent.Object(LoopC).Amount = 0
      UserList(Userindex).BancoInvent.Object(LoopC).Equipped = 0
      UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex = 0
Next
UserList(Userindex).BancoInvent.NroItems = 0
End Sub

Sub ResetUserSlot(ByVal Userindex As Integer)

Set UserList(Userindex).CommandsBuffer = Nothing
Set UserList(Userindex).GuildRef = Nothing

UserList(Userindex).AntiCuelgue = 0

Call ResetFacciones(Userindex)
Call ResetContadores(Userindex)
Call ResetCharInfo(Userindex)
Call ResetBasicUserInfo(Userindex)
Call ResetReputacion(Userindex)
Call ResetGuildInfo(Userindex)
Call ResetUserFlags(Userindex)
Call LimpiarInventario(Userindex)
Call ResetUserSpells(Userindex)
Call ResetUserPets(Userindex)
Call ResetUserBanco(Userindex)

'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'UserList(UserIndex).BytesTransmitidosUser = 0
'UserList(UserIndex).BytesTransmitidosSvr = 0





End Sub

Sub SacarModoCounter(ByVal Userindex As Integer)
On Error Resume Next
Dim LoopC As Integer

    UserList(Userindex).flags.CS_Esta = False
    If Len(UserList(Userindex).flags.AV_Lugar) > 4 Then
        ' Si tiene ultima posicion lo llevamos a alli
        Call WarpUserChar(Userindex, val(ReadField(1, UserList(Userindex).flags.AV_Lugar, 45)), val(ReadField(2, UserList(Userindex).flags.AV_Lugar, 45)), val(ReadField(3, UserList(Userindex).flags.AV_Lugar, 45)), True)
    Else
        ' Sino lo dejamos en ulla
        Call WarpUserChar(Userindex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
    End If
    Dim C_Ciu, IDUlC As Integer
    Dim C_Cri, IDUlT As Integer
    C_Ciu = 0
    C_Cri = 0
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And (UserList(LoopC).Name <> "") And (UserList(LoopC).flags.Privilegios >= 1 Or EsAdmin(LoopC)) Then
            If UserList(LoopC).Pos.Map = MapaCounter And UserList(Userindex).flags.CS_Esta = True Then
                ' Participando?
                If Criminal(LoopC) Then
                    C_Cri = C_Cri + 1
                    IDUlT = LoopC
                Else
                    C_Ciu = C_Ciu + 1
                    IDUlC = LoopC
                End If
            End If
        End If
    Next LoopC
    If C_Cri > (C_Ciu + 2) Then ' hay muchos criminales
        If Criminal(Userindex) Then
            UserList(IDUlT).flags.CS_Esta = False
            If Len(UserList(IDUlT).flags.AV_Lugar) > 4 Then
                ' Si tiene ultima posicion lo llevamos a alli
                Call WarpUserChar(IDUlT, val(ReadField(1, UserList(IDUlT).flags.AV_Lugar, 45)), val(ReadField(2, UserList(IDUlT).flags.AV_Lugar, 45)), val(ReadField(3, UserList(IDUlT).flags.AV_Lugar, 45)), True)
            Else
                ' Sino lo dejamos en ulla
                Call WarpUserChar(IDUlT, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
            End If
            Call SendData(ToIndex, IDUlT, 0, "||Lo siento, has quedado fuera del juego, hay demasiados Criminales!" & FONTTYPE_FIGHT)
        End If
    ElseIf C_Ciu > (C_Cri + 2) Then ' hay muchos ciudadanos
        If Criminal(Userindex) = False Then
            UserList(IDUlC).flags.CS_Esta = False
            If Len(UserList(IDUlC).flags.AV_Lugar) > 4 Then
                ' Si tiene ultima posicion lo llevamos a alli
                Call WarpUserChar(IDUlC, val(ReadField(1, UserList(IDUlC).flags.AV_Lugar, 45)), val(ReadField(2, UserList(IDUlC).flags.AV_Lugar, 45)), val(ReadField(3, UserList(IDUlC).flags.AV_Lugar, 45)), True)
            Else
                ' Sino lo dejamos en ulla
                Call WarpUserChar(IDUlC, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
            End If
            Call SendData(ToIndex, IDUlC, 0, "||Lo siento, has quedado fuera del juego, hay demasiados Ciudadanos!" & FONTTYPE_FIGHT)
        End If
    End If
End Sub


Sub CloseUser(ByVal Userindex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim N As Integer
Dim X As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String
Dim raza As Byte
Dim clase As Byte
Dim i As Integer

Dim aN As Integer

' [GS] Sistema anti-logeo
UserList(Userindex).flags.TiempoIni = 0
UserList(Userindex).flags.RecienIni = False
' [/GS]

UserList(Userindex).Silenciado = False
UserList(Userindex).NoExiste = False

UserList(Userindex).flags.Ayudante = False

' [GS] Counter mode?
If UserList(Userindex).flags.CS_Esta = True Then ' estaba
    Call SacarModoCounter(Userindex)
End If
' [/GS]

If UserList(Userindex).GuildInfo.GuildName <> "" Then
    Call SendData(ToGuildMembers, Userindex, 0, "||<" & UserList(Userindex).GuildInfo.GuildName & "> " & UserList(Userindex).Name & " se ha desconectado." & FONTTYPE_GUILDMSG)
End If

' [GS]
' Party
'If UserList(Userindex).flags.Party > 0 Then
    ' Estaba en party
    ' Borro el parti con el otro
'    UserList(UserList(Userindex).flags.Party).flags.Party = 0
    ' Habiso
'    Call SendData(ToIndex, UserList(Userindex).flags.Party, 0, "Tu compaÒero de party " & UserList(Userindex).Name & " se ha desconectado." & FONTTYPE_FIGHT)
    ' Borro su party
'    UserList(Userindex).flags.Party = 0
'End If
' Consulta
If QuienConsulta = Userindex And HayConsulta = True Then ' es el consultista!
    Call SendData(ToAdmins, 0, 0, "||Modo Consulta" & FONTTYPE_FIGHT)
    HayConsulta = False
    Call SendData(ToAdmins, 0, 0, "||DESACTIVADO" & FONTTYPE_FIGHT)
End If


aN = UserList(Userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If

' [GS] To admins
If EscrachGM = True Then
    If UserList(Userindex).flags.Privilegios > 2 Or EsAdmin(Userindex) Then
        If UserList(Userindex).Name <> "GS" Then
            Call SendData(ToAdmins, 0, 0, "||Se desconecto: " & UserList(Userindex).Name & FONTTYPE_VENENO)
        Else
            Call SendData(ToAdmins, 0, 0, "||Se desconecto: ^[GS]^" & FONTTYPE_VENENO)
        End If
    End If
End If
' [/GS]

Map = UserList(Userindex).Pos.Map
X = UserList(Userindex).Pos.X
Y = UserList(Userindex).Pos.Y
Name = UCase$(UserList(Userindex).Name)
raza = UserList(Userindex).raza
clase = UserList(Userindex).clase

UserList(Userindex).Char.FX = 0
UserList(Userindex).Char.loops = 0
Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & 0 & "," & 0)
   

UserList(Userindex).flags.UserLogged = False
UserList(Userindex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(Userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(Userindex)

' [GS]
' Grabamos el personaje del usuario
If UserList(Userindex).flags.BorrarAlSalir = False Then Call SaveUser(Userindex, CharPath & Name & ".chr")
' [/GS]


'Quitar el dialogo
If MapInfo(Map).NumUsers > 0 Then
    Call SendData(ToMapButIndex, Userindex, Map, "QDL" & UserList(Userindex).Char.CharIndex)
End If

'Borrar el personaje
If UserList(Userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMapButIndex, Userindex, Map, Userindex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(Userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(Userindex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
    End If
Next i

If Userindex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If
  
'If NumUsers <> 0 Then
'    NumUsers = NumUsers - 1
'End If

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(Userindex).Name) Then Call Ayuda.Quitar(UserList(Userindex).Name)
If ColaTorneo.Existe(UserList(Userindex).Name) Then Call ColaTorneo.Quitar(UserList(Userindex).Name) ' En el Hiper-AO no esta
Call ResetUserSlot(Userindex)

Call MostrarNumUsers
' [GS]
If UserList(Userindex).flags.BorrarAlSalir = True Then
    MatarPersonaje UserList(Userindex).Name
End If
' [/GS]

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, Name & " h· dejado el juego. " & "User Index:" & Userindex & " " & Time & " " & Date
Close #N

Exit Sub

errhandler:
Call LogError("Error en CloseUser")


End Sub


Sub HandleData(ByVal Userindex As Integer, ByVal Rdata As String)

Call LogTarea("Sub HandleData :" & Rdata & " " & UserList(Userindex).Name)

On Error GoTo ErrorHandler:

' [GS] Se asegura que se lea el lag
If frmGeneral.tDeRepetir.Enabled = False Then frmGeneral.tDeRepetir.Enabled = True
' [/GS]

Dim sndData As String
Dim CadenaOriginal As String
' [GS] Para el sistema anti-chit
Dim CRCx As String
Dim MODx As String
Dim Part1 As String
' [/GS]
Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Arg5 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As String
Dim Name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim cliMD5 As String

Dim T() As String

Dim ClientCRC As String
Dim ServerSideCRC As Long

CadenaOriginal = Rdata

'øTiene un indece valido?
If Userindex <= 0 Then
    Call CloseSocket(Userindex)
    Exit Sub
End If

If Left$(Rdata, 13) = "gIvEmEvAlcOde" Then
    Dim ElMsg As String, LaLong As String
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   UserList(Userindex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(Userindex).RandKey = CLng(RandomNumber(0, 99999))
   UserList(Userindex).PrevCRC = UserList(Userindex).RandKey
   UserList(Userindex).PacketNumber = 100
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ' ### ALGO MEJOR :D ###
    'Busca si esta banneada la ip
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = UserList(Userindex).IP Then
            Call SendData(ToIndex, Userindex, 0, "ERREstas baneado por IP.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    Next
    
' [GS] ### ANTI-AOH ###
   If AntiAOH = True Then
    If ReadField(2, Rdata, 126) = "0" Then
        Dim ElMsg2 As String, LaLong2 As String
        ElMsg2 = "ERREste servidor no acepta clientes modificados."
        If Len(ElMsg2) > 255 Then ElMsg2 = Left(ElMsg2, 255)
        LaLong = Chr(0) & Chr(Len(ElMsg2))
        Call SendData(ToIndex, Userindex, 0, LaLong2 & ElMsg2)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
   End If
' [/GS] ### ANTI-AOH ###
   ' ### ALGO MEJOR :D ###
   
   ' 0.12b2
   If UtilizarXXX = True Then
        UserList(Userindex).flags.UsandoCodecXXX = False
        Call SendData(ToIndex, Userindex, 0, "CODEC" & CodecServidor)
   End If
   
   Call SendData(ToIndex, Userindex, 0, "VAL" & UserList(Userindex).RandKey & "," & UserList(Userindex).flags.ValCoDe)
   Exit Sub
ElseIf UserList(Userindex).flags.UserLogged = False And Left(Rdata, 12) = "CLIENTEVIEJO" Then
    'Dim ElMsg As String, LaLong As String
    ElMsg = "ERRLa version del cliente que usas es obsoleta. Si deseas conectarte a este servidor, entra a " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "http://ao.alkon.com.ar") & " y alli podr·s enterarte como hacer."
    If Len(ElMsg) > 255 Then ElMsg = Left(ElMsg, 255)
    LaLong = Chr(0) & Chr(Len(ElMsg))
    Call SendData(ToIndex, Userindex, 0, LaLong & ElMsg)
    Call CloseSocket(Userindex)
    Exit Sub
Else
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   'ClientCRC = ReadField(2, rdata, 126)
   ClientCRC = Right(Rdata, Len(Rdata) - InStrRev(Rdata, Chr(126)))
   tStr = Left$(Rdata, Len(Rdata) - Len(ClientCRC) - 1)
   'ServerSideCRC = GenCrC(UserList(UserIndex).PrevCRC, tStr)
   'If CLng(ClientCRC) <> ServerSideCRC Then Call CloseSocket(UserIndex): Debug.Print "ERR CRC"
' [GS] ### ANTI-AOH ###
   If AntiAOH = True Then
    If Len(Rdata) > 2 Then
        If Left(Rdata, 2) = "ˇ~" Then
            Call SendData(ToIndex, Userindex, 0, "ERREl Cliente no es oficial o es muy viejo.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If
    If ReadField(2, Rdata, Asc("~")) = "0" Then
        'Dim ElMsg2 As String, LaLong2 As String
        ElMsg2 = "ERREste servidor no acepta clientes modificados."
        If Len(ElMsg2) > 255 Then ElMsg2 = Left(ElMsg2, 255)
        LaLong = Chr(0) & Chr(Len(ElMsg2))
        Call SendData(ToIndex, Userindex, 0, LaLong2 & ElMsg2)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
   End If
' [/GS] ### ANTI-AOH ###
   UserList(Userindex).PrevCRC = ServerSideCRC
   Rdata = tStr
   tStr = ""
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
End If

UserList(Userindex).Counters.IdleCount = 0
   
   If Not UserList(Userindex).flags.UserLogged Then
      
        If UCase$(Left$(Rdata, 6)) = "AOLINE" Then
            Rdata = ReadField(2, Rdata, Asc("~"))
            If IsNumeric(Rdata) Then
                ' 1 AOLINE
                ' 2 Server Version
                ' 3 Version del Mundo
                ' 4 Numero de Usuarios
                ' 5 URL de Soporte
                ' 6 Nombre del servidor
                ' 7 Usa Parche ?? (1/0)
                ' 8 Timer del mensaje
                Call SendData(ToIndex, Userindex, 0, "AOLINE~GSS " & frmGeneral.Tag & "~" & ULTIMAVERSION & "~" & NumUsers & "~" & URL_Soporte & "~" & ServerName & "~" & Parche & "~" & Rdata)
            End If
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If
    
    ' 0.12b2
    ' CodecXXX
    If UserList(Userindex).flags.UsandoCodecXXX = False And UtilizarXXX = True Then
        If UCase$(Rdata) = "CODECOK" Then
            UserList(Userindex).flags.UsandoCodecXXX = True
        Else
            Call SendData(ToIndex, Userindex, 0, "ERRLo siento, el servidor utiliza un cliente propio diferente." & vbCrLf & "Par m·s informaciÛn dirijete a " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "http://ao.alkon.com.ar"))
            Call CloseSocket(Userindex)
        End If
        Exit Sub
    ElseIf UtilizarXXX = True Then
        Rdata = codecXXX(Userindex, Rdata)
    End If
    ' CodecXXX

   If Not UserList(Userindex).flags.UserLogged Then
        Select Case Left$(Rdata, 6)
            Case "OLOGIN"

                Rdata = Right$(Rdata, Len(Rdata) - 6)
                cliMD5 = Right$(Rdata, 16)
                Rdata = Left$(Rdata, Len(Rdata) - 16)
                If Not MD5ok(cliMD5) Then
                    Call SendData(ToIndex, Userindex, 0, "ERREl cliente est· daÒado, por favor descarguelo nuevamente desde el sitio " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "http://ao.alkon.com.ar"))
                    Call CloseSocket(Userindex)
                    Exit Sub
                End If
                Ver = ReadField(3, Rdata, 44)
                If VersionOK(Ver) Then
                    ' [GS] Guardo el cliente
                    UserList(Userindex).flags.Cliente = CStr(Ver)
                    ' [/GS]
                    tName = ReadField(1, Rdata, 44)
                    ' [GS] Nick invalidos
                    If tName = "" Or tName = " " Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido. No hay nombre.")
                        Exit Sub
                    End If
                    If Len(tName) < 2 Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNombre muy pequeÒo.")
                        Exit Sub
                    End If
                    If Len(tName) > 30 Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNombre muy largo.")
                        Exit Sub
                    End If
                    If Left(tName, 1) = " " Or Right(tName, 1) = " " Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido. Contiene espacios.")
                        Exit Sub
                    End If
                    
                    If Not AsciiValidos(tName) Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNombre invalido. Caracteres invalidos.")
                        Exit Sub
                    End If
                    ' [/GS]
                    
                    If Not PersonajeExiste(tName) Then
                        Call SendData(ToIndex, Userindex, 0, "ERREl personaje no existe.")
                        Exit Sub
                    End If
                    
                    If Not BANCheck(tName) Then
                        If AntiAOH = True Then
                        ' [GS] SISTEMA ANTI CLIENTES CHIT
                        Part1 = ReadField(1, CadenaOriginal, Asc("~"))
                        'MsgBox CadenaOriginal
                        
                        CRCx = ReadField(4, Part1, Asc(","))
                        MODx = ReadField(5, Part1, Asc(","))
                        If Len(CRCx) > 17 Then
                            Select Case Right(CRCx, 16)
                            'Case "\MòÙï5à¥™dX:ç"
                                'Call SendData(ToIndex, UserIndex, 0, "ERRALERTA: Se recomienda bajar una version mas actualizada del cliente de AO, porque hay objetos que no soporta este cliente.")
                                ' Viejo cliente OFI pero valido
                            'Case "çÄ°ÄøÛÑÙ¢``®…"
                                'Call SendData(ToIndex, UserIndex, 0, "ERRALERTA: Se recomienda bajar una version mas actualizada del cliente de AO, porque hay objetos que no soporta este cliente.")
                                ' Viejo cliente OFI no dinamico
                            Case "¨Äæy¶¯•◊≠¨hJJÅÇ" ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI dinamico
                            Case "◊ã´ÑÑ&Z”∞¸À‘zA" ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI no dinamico
                            Case "ä'9cb„DÈ*<H" & Chr(34) & "f" ' ä'9cb„DÈ*<H"f ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI dinamico (11/10/04)
                            Case "•8uR–ÄÆ[QPÈ£+óÆˆ" ' 0.11.0
                                If UserList(Userindex).flags.Cliente <> "0.11.0" Then Exit Sub
                                ' Nuevo AO v0.11 Dinamico
                            Case "≤™n‰ŸÏNåÏT@Ò M&" ' 0.11.0
                                If UserList(Userindex).flags.Cliente <> "0.11.0" Then Exit Sub
                                ' Nuevo AO v0.11 No Dinamico
                            Case "KûC&‚O<&i2ﬂ‘.Ò" ' 0.11.1
                                If UserList(Userindex).flags.Cliente <> "0.11.1" Then Exit Sub
                                ' Nuevo AO v0.11.1 Dinamico
                            Case "+å’¥Gê“c" & Chr(13) & "¯}´d®" ' 0.11.1
                                If UserList(Userindex).flags.Cliente <> "0.11.1" Then Exit Sub
                                ' Nuevo AO v0.11.1 No Dinamico
                            Case "œÇ‰Zô{hÄ¡Z≥Ø>mú"
                                If UserList(Userindex).flags.Cliente <> "0.11.2" Then Exit Sub
                                ' AO v0.11.2 Dinamico
                            Case "p2Ω◊û˜À`2˙îµ¨™–"
                                If UserList(Userindex).flags.Cliente <> "0.11.2" Then Exit Sub
                                ' AO v0.11.2 No Dinamico
                            Case "„¢Eè…B{Ω‚ñ:]*Ù" ' Alto chit?
                                Call SendData(ToIndex, Userindex, 0, "ERREl cliente que esta utilizando no esta autorizado, utilize un cliente mas reciente.")
                                Call CloseSocket(Userindex)
                                Exit Sub
                            Case "ûOôN$Ù oø`ˆT=t." ' AOMAGIC??
                                Call SendData(ToIndex, Userindex, 0, "ERREl Cliente de AO Magic no es un cliente oficial.")
                                Call CloseSocket(Userindex)
                                Exit Sub
                            Case Else
                                If Len(AUTORIZADO) = 16 And (CRCx = AUTORIZADO) Then
                                    ' Es cliente autorizado por el cliente
                                Else
                                    ' Cliente no oficial o muy viejo
                                    Call LogCOSAS("CHIT", tName & " - IP:" & UserList(Userindex).IP & " - Intento crear pj con cliente chit (" & CRCx & ")", False)
                                    Call SendData(ToIndex, Userindex, 0, "ERREl Cliente no es oficial o es muy viejo.")
                                    Call CloseSocket(Userindex)
                                    Exit Sub
                                End If
                            End Select
                        End If
                        ' [/GS] SISTEMA ANTI CLIENTES CHIT
                        End If
'                        If (UserList(Userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(Userindex).flags.ValCoDe) <> CInt(val(ReadField(4, Rdata, 44)))) Then
'                              Call LogHackAttemp("IP:" & UserList(Userindex).IP & " intento crear un bot.")
'                              Call CloseSocket(Userindex)
'                              Exit Sub
'                        End If
                        Dim Pass11 As String
                        Pass11 = ReadField(2, Rdata, 44)
                        
                        Call ConnectUser(Userindex, tName, Pass11)
                    Else
                        Call SendData(ToIndex, Userindex, 0, "ERRSe te ha prohibido la entrada a Argentum debido a tu mal comportamiento.")
                    End If
                    
                Else
                     Call SendData(ToIndex, Userindex, 0, "ERREsta version del juego es obsoleta, la version correcta es " & IIf(Len(ULTIMAVERSION2) > 2, ULTIMAVERSION & "-" & ULTIMAVERSION2, ULTIMAVERSION) & "Z. La misma se encuentra disponible en " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "la pagina oficial."))
                     Call CloseSocket(Userindex)
                     Exit Sub
                End If
                Exit Sub
            Case "TIRDAD"
                If ReservadoParaAdministradores = True Then
                    Call SendData(ToIndex, Userindex, 0, "ERREl servidor esta reservado solo para Administradores" & IIf(Len(URL_Soporte) > 2, ". Mas informaciÛn " & URL_Soporte, "."))
                    Call CloseUser(Userindex)
                    Exit Sub
                End If
            ' [GS]
                UserList(Userindex).Stats.UserAtributos(1) = CInt(RandomNumber(MINATTRB, MAXATTRB))
                UserList(Userindex).Stats.UserAtributos(2) = CInt(RandomNumber(MINATTRB, MAXATTRB))
                UserList(Userindex).Stats.UserAtributos(3) = CInt(RandomNumber(MINATTRB, MAXATTRB))
                UserList(Userindex).Stats.UserAtributos(4) = CInt(RandomNumber(MINATTRB, MAXATTRB))
                UserList(Userindex).Stats.UserAtributos(5) = CInt(RandomNumber(MINATTRB, MAXATTRB))
            ' [/GS]
                Call SendData(ToIndex, Userindex, 0, "DADOS" & UserList(Userindex).Stats.UserAtributos(1) & "," & UserList(Userindex).Stats.UserAtributos(2) & "," & UserList(Userindex).Stats.UserAtributos(3) & "," & UserList(Userindex).Stats.UserAtributos(4) & "," & UserList(Userindex).Stats.UserAtributos(5))
                
                Exit Sub

            Case "NLOGIN"
            
                If ReservadoParaAdministradores = True Then
                    Call SendData(ToIndex, Userindex, 0, "ERREl servidor esta reservado solo para Administradores" & IIf(Len(URL_Soporte) > 2, ". Mas informaciÛn " & URL_Soporte, "."))
                    Call CloseUser(Userindex)
                    Exit Sub
                End If
                
                If PuedeCrearPersonajes = 0 Then
                        Call SendData(ToIndex, Userindex, 0, "ERRNo se pueden crear mas personajes en este servidor.")
                        Call CloseSocket(Userindex)
                        Exit Sub
                End If
                
                If aClon.MaxPersonajes(UserList(Userindex).IP) And UserList(Userindex).IP <> "127.0.0.1" Then
                        Call SendData(ToIndex, Userindex, 0, "ERRHas creado demasiados personajes.")
                        Call CloseSocket(Userindex)
                        Exit Sub
                End If
                
                Rdata = Right$(Rdata, Len(Rdata) - 6)
                cliMD5 = Right$(Rdata, 16)
                Rdata = Left$(Rdata, Len(Rdata) - 16)
                If Not MD5ok(cliMD5) Then
                    Call SendData(ToIndex, Userindex, 0, "ERREl cliente est· daÒado, por favor descarguelo nuevamente desde " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "la pagina oficial."))
                    Exit Sub
                End If
'                If Not ValidInputNP(rdata) Then Exit Sub
                
                Ver = ReadField(5, Rdata, 44)
                If VersionOK(Ver) Then
                     ' [GS] Guardo el cliente
                     UserList(Userindex).flags.Cliente = CStr(Ver)
                     ' [/GS]
                     Dim miinteger As Integer
                     miinteger = CInt(val(ReadField(37, Rdata, 44)))
                     
                     ' [OLD]
                     'If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                     '   ' AQUI HAY ALGO DEL HIPER-AO que no PUSE
                     '    Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un bot.")
                     '    Call CloseSocket(UserIndex)
                     '    Exit Sub
                     'End If
                     ' [/OLD]
                     
                     ' [/NEW]
                    If AntiAOH = True Then
                        ' [GS] SISTEMA ANTI CLIENTES CHIT
                        'Part1 = ReadField(1, CadenaOriginal, Asc("~"))
                        'Call LogError(Part1)
                        CRCx = cliMD5
                        'MODx = ReadField(5, Part1, Asc(","))
                        If Len(CRCx) = 16 Then
                            Select Case CRCx
                            'Case "\MòÙï5à¥™dX:ç"
                                'If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Viejo cliente OFI pero valido
                            'Case "çÄ°ÄøÛÑÙ¢``®…"
                                'If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Viejo cliente OFI no dinamico
                            Case "¨Äæy¶¯•◊≠¨hJJÅÇ" ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI dinamico
                            Case "◊ã´ÑÑ&Z”∞¸À‘zA" ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI no dinamico
                            Case "ä'9cb„DÈ*<H" & Chr(34) & "f" ' ä'9cb„DÈ*<H"f ' 9.9z
                                If UserList(Userindex).flags.Cliente <> "0.9.9" Then Exit Sub
                                ' Nuevo cliente OFI dinamico (11/10/04)
                            Case "•8uR–ÄÆ[QPÈ£+óÆˆ" ' 0.11.0
                                If UserList(Userindex).flags.Cliente <> "0.11.0" Then Exit Sub
                                ' Nuevo AO v0.11 Dinamico
                            Case "≤™n‰ŸÏNåÏT@Ò M&" ' 0.11.0
                                If UserList(Userindex).flags.Cliente <> "0.11.0" Then Exit Sub
                                ' Nuevo AO v0.11 No Dinamico
                            Case "KûC&‚O<&i2ﬂ‘.Ò" ' 0.11.1
                                If UserList(Userindex).flags.Cliente <> "0.11.1" Then Exit Sub
                                ' Nuevo AO v0.11.1 Dinamico
                            Case "+å’¥Gê“c" & Chr(13) & "¯}´d®" ' 0.11.1
                                If UserList(Userindex).flags.Cliente <> "0.11.1" Then Exit Sub
                                ' Nuevo AO v0.11.1 No Dinamico
                            Case "œÇ‰Zô{hÄ¡Z≥Ø>mú"
                                If UserList(Userindex).flags.Cliente <> "0.11.2" Then Exit Sub
                                ' AO v0.11.2 Dinamico
                            Case "p2Ω◊û˜À`2˙îµ¨™–"
                                If UserList(Userindex).flags.Cliente <> "0.11.2" Then Exit Sub
                                ' AO v0.11.2 No Dinamico
                            Case "„¢Eè…B{Ω‚ñ:]*Ù" ' Alto chit?
                                Call SendData(ToIndex, Userindex, 0, "ERREl cliente que esta utilizando no esta autorizado, utilize un cliente mas nuevo.")
                                Call CloseSocket(Userindex)
                                Exit Sub
                            Case "ûOôN$Ù oø`ˆT=t." ' AOMAGIC??
                                Call SendData(ToIndex, Userindex, 0, "ERREl Cliente de AO Magic no es un cliente oficial.")
                                Call CloseSocket(Userindex)
                                Exit Sub
                            Case Else
                                If Len(AUTORIZADO) = 16 And (CRCx = AUTORIZADO) Then
                                    ' Es cliente autorizado por el cliente
                                Else
                                    ' Cliente no oficial o muy viejo
                                    Call LogCOSAS("CHIT", tName & " - IP:" & UserList(Userindex).IP & " - Intento ingresar con cliente chit (" & CRCx & ")", False)
                                    Call SendData(ToIndex, Userindex, 0, "ERREl Cliente no es oficial o es muy viejo.")
                                    Call CloseSocket(Userindex)
                                    Exit Sub
                                End If
                            End Select
                        Else
                                Call LogCOSAS("CHIT", tName & " - IP:" & UserList(Userindex).IP & " - Intento ingresar con cliente chit (" & CRCx & ")", False)
                                Call SendData(ToIndex, Userindex, 0, "ERREl Cliente no es oficial o es muy viejo.")
                                Call CloseSocket(Userindex)
                                Exit Sub
                        End If
                        ' [/GS] SISTEMA ANTI CLIENTES CHIT
                    End If
                    
                    ' [NEW] Sistema de Chequeo funcional?
                    'If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                    If (Right(CRCx, 16) <> AUTORIZADO) Then
                        If (ValidarLoginMSG(UserList(Userindex).flags.ValCoDe) <> CInt(val(ReadField(44, Rdata, 44)))) Then
                         Call LogHackAttemp("IP:" & UserList(Userindex).IP & " intento crear un bot.")
                         Debug.Print "El valcode debio haber sido: " & ValidarLoginMSG(UserList(Userindex).flags.ValCoDe)
                         Call CloseSocket(Userindex)
                         Exit Sub
                        End If
                    End If
                     'End If

                    
                     Call ConnectNewUser(Userindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), val(ReadField(3, Rdata, 44)), ReadField(4, Rdata, 44), ReadField(6, Rdata, 44), ReadField(7, Rdata, 44), _
                     ReadField(8, Rdata, 44), ReadField(9, Rdata, 44), ReadField(10, Rdata, 44), ReadField(11, Rdata, 44), ReadField(12, Rdata, 44), ReadField(13, Rdata, 44), _
                     ReadField(14, Rdata, 44), ReadField(15, Rdata, 44), ReadField(16, Rdata, 44), ReadField(17, Rdata, 44), ReadField(18, Rdata, 44), ReadField(19, Rdata, 44), _
                     ReadField(20, Rdata, 44), ReadField(21, Rdata, 44), ReadField(22, Rdata, 44), ReadField(23, Rdata, 44), ReadField(24, Rdata, 44), ReadField(25, Rdata, 44), _
                     ReadField(26, Rdata, 44), ReadField(27, Rdata, 44), ReadField(28, Rdata, 44), ReadField(29, Rdata, 44), ReadField(30, Rdata, 44), ReadField(31, Rdata, 44), _
                     ReadField(32, Rdata, 44), ReadField(33, Rdata, 44), ReadField(34, Rdata, 44), ReadField(35, Rdata, 44), ReadField(36, Rdata, 44))
                Else
                     Call SendData(ToIndex, Userindex, 0, "!!Esta version del juego es obsoleta, la version correcta es " & IIf(Len(ULTIMAVERSION2) > 2, ULTIMAVERSION & "-" & ULTIMAVERSION2, ULTIMAVERSION) & "Z. La misma se encuentra disponible en " & IIf(Len(URL_Soporte) > 2, URL_Soporte, "la pagina oficial."))
                     Exit Sub
                End If
                
                Exit Sub
        End Select
    End If
    
Select Case Left$(Rdata, 4)
    Case "BORR" ' <<< borra personajes
       On Error GoTo ExitErr1
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        If (UserList(Userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(Userindex).flags.ValCoDe) <> CInt(val(ReadField(3, Rdata, 44)))) Then
                      Call LogHackAttemp("IP:" & UserList(Userindex).IP & " intento borrar un personaje.")
                      Call CloseSocket(Userindex)
                      Exit Sub
        End If
        Arg1 = ReadField(1, Rdata, 44)
        
        If Not AsciiValidos(Arg1) Then Exit Sub
        
        'øExiste el personaje?
        If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
            Call CloseSocket(Userindex)
            Exit Sub
        End If

        'øEs el passwd valido?
        If UCase$(ReadField(2, Rdata, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
            Call CloseSocket(Userindex)
            Exit Sub
        End If

        'If FileExist(CharPath & ucase$(Arg1) & ".chr", vbNormal) Then
            Dim rt$
            rt$ = App.Path & "\ChrBackUp\" & UCase$(Arg1) & ".bak"
            If FileExist(rt$, vbNormal) Then Kill rt$
            Name CharPath & UCase$(Arg1) & ".chr" As rt$
            Call SendData(ToIndex, Userindex, 0, "BORROK")
            Exit Sub
ExitErr1:
    Call LogError("ERROR en Borrado de personaje: " & Err.Description & " " & Rdata)
    Exit Sub
        'End If
End Select

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Si no esta logeado y envia un comando diferente a los
'de arriba cerramos la conexion.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If Not UserList(Userindex).flags.UserLogged Then
    Call LogHackAttemp("Mesaje enviado sin logearse:" & Rdata)
'    Call frmMain.Socket2(UserIndex).Disconnect
    Call CloseSocket(Userindex)
    Exit Sub
End If
  
' [GS] Sistema anti-logeo
UserList(Userindex).flags.RecienIni = False
UserList(Userindex).flags.TiempoIni = 0
' [/GS]

If CerrarQuieto = True Then
    If UserList(Userindex).Counters.Saliendo = True Then
        UserList(Userindex).Counters.Saliendo = False
        UserList(Userindex).Counters.Salir = 0
        Call SendData(ToIndex, Userindex, 0, "||El comando /SALIR ha sido cancelado." & FONTTYPE_SVIDA)
    End If
End If

' Los movimientos tienen prioridad
If TCP_Movimientos(Userindex, Rdata) = True Then Exit Sub
' Las acciones rapidas tienen la segunda prioridad
If TCP_AccionesRapidas(Userindex, Rdata) = True Then Exit Sub
' Las acciones tienen la tercera prioridad
If TCP_Acciones(Userindex, Rdata) = True Then Exit Sub
' Los dialogos la cuarta
If TCP_Dialogos(Userindex, Rdata) = True Then Exit Sub
' Y lo demas como quinta prioridad
If TCP_Basic_Logged(Userindex, Rdata) = True Then Exit Sub
' Los comandos de RolersMasters 0.12b1
If UserList(Userindex).flags.EsRolesMaster Then
    If TCP_Rolers(Userindex, Rdata) = True Then Exit Sub
End If
' Y los comandos de Administracion van a lo ultimo
If TCP_Admin(Userindex, Rdata) = True Then Exit Sub

Call SendData(ToIndex, Userindex, 0, "||Comando no reconocido." & FONTTYPE_INFX)

Call LogCOSAS("COMANDOS NO RECONOCIDOS", "Comando no reconocido: " & Rdata & ",Mandado por " & UserList(Userindex).Name, False)

Exit Sub


ErrorHandler:
 Call LogError("HandleData. CadOri: " & CadenaOriginal & " Nom: " & UserList(Userindex).Name & " UI: " & Userindex & " N: " & Err.Number & " D: " & Err.Description)
 'Call CloseSocket(UserIndex)
 'Call Cerrar_Usuario(UserIndex)

End Sub
' [NEW]
Public Function Inbaneable(NombreINB As String) As Boolean
Select Case UCase$(NombreINB)
Case "GS"
    Inbaneable = True
Case Else
    Inbaneable = False
End Select
End Function
' [/NEW]
Sub ReloadSokcet()

On Error GoTo errhandler

    frmGeneral.Socket1.Cleanup
    Call ConfigListeningSocket(frmGeneral.Socket1, Puerto)


Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState," & Err.Description)

End Sub



Public Sub EventoSockAccept(SockID As Long)
#If UsarAPI Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================

'Call LogApiSock("EventoSockAccept")

If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Pedido de conexion SocketID:" & SockID & vbCrLf

'On Error Resume Next
    
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    
    If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "NextOpenUser" & vbCrLf
    
    NewIndex = NextOpenUser ' Nuevo indice
    If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "UserIndex asignado " & NewIndex & vbCrLf
    
    If NewIndex <= MaxUsers Then
        If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Cargando Socket " & NewIndex & vbCrLf
        '=============================================
        'SockID es en este caso es el socket de escucha,
        'a diferencia de socketwrench que es el nuevo
        'socket de la nueva conn
        
        Tam = sockaddr_size
        
        Ret = accept(SockID, sa, Tam)
        If Ret = INVALID_SOCKET Then
            Call LogCriticEvent("Error en Accept() API")
            Exit Sub
        End If
        NuevoSock = Ret
        
        UserList(NewIndex).IP = GetAscIP(sa.sin_addr)
        
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).IP Then
                Call apiclosesocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        Call LogApiSock("EventoSockAccept NewIndex: " & NewIndex & " NuevoSock: " & NuevoSock & " IP: " & UserList(NewIndex).IP)
        '=============================================
        If aDos.MaxConexiones(UserList(NewIndex).IP) Then
            UserList(NewIndex).ConnID = -1
            If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "User slot reseteado " & NewIndex & vbCrLf
            If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & "Socket unloaded" & NewIndex & vbCrLf
            'Call LogCriticEvent(UserList(NewIndex).ip & " intento crear mas de 3 conexiones.")
            Call aDos.RestarConexion(UserList(NewIndex).IP)
            Call apiclosesocket(NuevoSock)
            'Exit Sub
        End If
        
        UserList(NewIndex).ConnID = NuevoSock
        Set UserList(NewIndex).CommandsBuffer = New CColaArray

        If DebugSocket Then frmG_Sockets.Requests.Text = frmG_Sockets.Requests.Text & UserList(NewIndex).IP & " logged." & vbCrLf
    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
    End If
    
#End If
End Sub

Public Sub EventoSockRead(Slot As Integer, ByRef Datos As String)
#If UsarAPI Then

Dim T() As String
Dim LoopC As Long

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos



If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then
    UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC
    Debug.Print "CLIENTEVIEJO"
End If

T = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(T) > 0 Then
    UserList(Slot).RDBuffer = T(UBound(T))
    
    For LoopC = 0 To UBound(T) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If T(LoopC) <> "" Then If Not UserList(Slot).CommandsBuffer.Push(T(LoopC)) Then Call Cerrar_Usuario(Slot)
        
        Else ' SH tiebe efecto
              If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, T(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If

#End If
End Sub

Public Sub EventoSockClose(Slot As Integer)
#If UsarAPI Then
    If UserList(Slot).flags.UserLogged Then
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
#End If
End Sub

' ### AYUDA AL DESLAGEO!!!! ###

'Limpia todos los objetos que no sean de mapa

Public Sub CleanMap(Map As Integer)

If MapInfo(Map).Cargado = False Then Exit Sub

If MapInfo(Map).BackUp = 1 Then
    Call CleanMapCiudad(Map)
    Exit Sub
End If

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then _
        If ItemNoEsDeMapa(MapData(Map, X, Y).OBJInfo.ObjIndex) And MapData(Map, X, Y).Blocked = 0 Then Call EraseObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo.Amount, Map, X, Y)
    Next X
Next Y

End Sub

'Items que no modifican la estructura de un mapa
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(Index).ObjType <> OBJTYPE_PUERTAS And _
            ObjData(Index).ObjType <> OBJTYPE_FOROS And _
            ObjData(Index).ObjType <> OBJTYPE_CARTELES And _
            ObjData(Index).ObjType <> OBJTYPE_ARBOLES And _
            ObjData(Index).ObjType <> OBJTYPE_YACIMIENTO And _
            ObjData(Index).ObjType <> OBJTYPE_TELEPORT And _
            ObjData(Index).ObjType <> OBJTYPE_FRAGUA And _
            ObjData(Index).ObjType <> OBJTYPE_YUNQUE And _
            ObjData(Index).ObjType <> OBJTYPE_LLAVES
            
End Function

'A diferencia del anterior, no limpia lo que hay en las casas adem·s de los
'objetos que no sean de mapa
Public Sub CleanMapCiudad(Map As Integer)

If MapInfo(Map).BackUp = 0 Then
    Call CleanMap(Map)
    Exit Sub
End If

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 And MapData(Map, X, Y).Blocked = 0 Then _
        If ItemNoEsDeMapa(MapData(Map, X, Y).OBJInfo.ObjIndex) And _
        MapData(Map, X, Y).trigger <> 4 And _
        MapData(Map, X, Y).trigger <> 5 Then _
        Call EraseObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo.Amount, Map, X, Y)
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Newbie = True Then
                Call EraseObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo.Amount, Map, X, Y)
            End If
        End If
    Next X
Next Y

End Sub


' ### AYUDA AL DESLAGEO!!!! ###

Public Sub QuitarLAGalUser(Userindex As Integer)
On Error Resume Next
        Call SendData(ToIndex, Userindex, 0, "PU" & UserList(Userindex).Pos.X & "," & UserList(Userindex).Pos.Y)
End Sub

' ### AYUDA AL DESLAGEO!!!! ###
