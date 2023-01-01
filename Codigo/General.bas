Attribute VB_Name = "General"
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


Global LeerNPCs As New clsLeerInis
Global LeerNPCsHostiles As New clsLeerInis

'[Sicarul] Hiper-AO
Public Entrando As Boolean
Public Cago As Boolean
Public CorrectPass As String
Public CorrectUser As String
'[/Sicarul]

Global ANpc As Long
Global Anpc_host As Long

Option Explicit

' [GS] Buscar actualizaciones automaticas
Public Function ActualizarAutomatico() As Boolean
On Error Resume Next
'Dim Datos As String
'frmGeneral.Inet1.RequestTimeout = 10
' Nueva URL : http://www.geocities.com/triforce_w/update.txt
'frmGeneral.Inet1.URL = Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(119) & Chr(119) & Chr(119) & Chr(46) & Chr(103) & Chr(101) & Chr(111) & Chr(99) & Chr(105) & Chr(116) & Chr(105) & Chr(101) & Chr(115) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(116) & Chr(114) & Chr(105) & Chr(102) & Chr(111) & Chr(114) & Chr(99) & Chr(101) & Chr(95) & Chr(119) & Chr(47) & Chr(117) & Chr(112) & Chr(100) & Chr(97) & Chr(116) & Chr(101) & Chr(46) & Chr(116) & Chr(120) & Chr(116)
'frmGeneral.Inet1.protocol = icHTTP ' cambiar si es otro protocolo(FTP:\\) (Https:\\)
'DoEvents
'Datos = frmGeneral.Inet1.OpenURL 'Accedemos al archivo :D
'DoEvents
'If Len(Datos) < 1 Then
'    ActualizarAutomatico = False
'    NuevaVersion = ""
'Else
'    NuevaVersion = ""
'    NuevaVersion = ReadField(2, Datos, Asc("¾")) 'dhowqwd¾(     )
'    If Len(NuevaVersion) < 1 Then
'        ActualizarAutomatico = False
'        Exit Function
'    End If
'    NuevaVersion = ReadField(1, NuevaVersion, Asc("¶")) ' ([   ]¶h5yt)
'    If Len(NuevaVersion) > 1 Then
'        ActualizarAutomatico = True
'    Else
'        ActualizarAutomatico = False
'    End If
'End If
End Function
' [/GS]


Sub DarCuerpoDesnudo(ByVal Userindex As Integer)

Select Case UserList(Userindex).raza
    Case RAZA_HUMANO
      Select Case UserList(Userindex).genero
                Case HOMBRE
                     UserList(Userindex).Char.Body = 21
                Case MUJER
                     UserList(Userindex).Char.Body = 39
      End Select
    Case RAZA_ELFO_OSCURO
      Select Case UserList(Userindex).genero
                Case HOMBRE
                     UserList(Userindex).Char.Body = 32
                Case MUJER
                     UserList(Userindex).Char.Body = 40
      End Select
    Case RAZA_ENANO
      Select Case UserList(Userindex).genero
                Case HOMBRE
                     UserList(Userindex).Char.Body = 53
                Case MUJER
                     UserList(Userindex).Char.Body = 60
      End Select
    Case RAZA_GNOMO
      Select Case UserList(Userindex).genero
                Case HOMBRE
                     UserList(Userindex).Char.Body = 53
                Case MUJER
                     UserList(Userindex).Char.Body = 60
      End Select
    Case Else
      Select Case UserList(Userindex).genero
                Case HOMBRE
                     UserList(Userindex).Char.Body = 21
                Case MUJER
                     UserList(Userindex).Char.Body = 39
      End Select
    
End Select

UserList(Userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)



End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function




Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer

For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i



End Sub

Sub EnviarSpawnList(ByVal Userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(ToIndex, Userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If Not (UsarAPI = 1) Then

Obj.AddressFamily = AF_INET
Obj.protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

If LCase$(Command) = "-mantenimiento" Then
    frmMantenimiento.Show
    frmMantenimiento.SEG.Caption = 60
    frmMantenimiento.MAN.Enabled = True
    frmMantenimiento.SetFocus
    Exit Sub
End If

' [GS] v0.14a1
If LCase$(Command) = "-config" Then
    If App.PrevInstance = True Then
        MsgBox "Otra aplicación relacionada con el Servidor se encuentra funcionando, por favor cierrela si desea entrar a Configurar unicamente."
        End
    End If
    If FileExist(App.Path & "\Opciones.ini", vbArchive) = False Then
        If FileExist(App.Path & "\Opciones.ini.default", vbArchive) = False Then
            MsgBox "ALERTA: El servidor no se encuentra configurado. Falta: opciones.ini.default", vbCritical
        Else
            Call FileCopy(App.Path & "\Opciones.ini.default", App.Path & "\Opciones.ini")
        End If
    End If
    frmCargando.Show
    frmCargando.Label1(0).Caption = "Cargando Opciones..."
    Call LoadSini
    Call LoadOpcsINI
    Unload frmCargando
    frmG_T_OPCIONES.Caption = "GSS >> Configuración || Opciones.ini (Versión Rapida)"
    frmG_T_OPCIONES.INISVR.Value = 1
    frmG_T_OPCIONES.INISVR.Visible = True
    frmG_T_OPCIONES.Tag = "UNICO"
    frmG_T_OPCIONES.Show
    Exit Sub
End If
' [/GS] v0.14a1

' [GS] Anti multi ejecucion
If App.PrevInstance And Command <> "-ejecutarigual" Then
    MsgBox "El servidor ya está siendo ejecutado."
    End
End If
' [/GS]
Dim nfile As Integer
Dim CarpBugs As String
CarpBugs = App.Path & "\REPORTES DE ERRORES\" & Replace(Replace(Date, "/", "-"), "\", "-") & "_" & Replace(Format(Time, "HH:mm"), ":", "-")
frmCargando.Show
frmGeneral.Visible = False

'Call LoadOpcsINI
'End

frmCargando.Label1(2).Caption = "Verificando autorización..."
DoEvents

If FileExist(App.Path & "\logs\Tiempos.log", vbNormal) Then BorrarArchivo (App.Path & "\logs\Tiempos.log")

' [GS] Validacion oficial
'"http://localhost/client.txt") = False Then '
GoTo NextVal:
If SoyValido(frmGeneral.Inet1, Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(99) & Chr(46) & Chr(49) & Chr(97) & Chr(115) & Chr(112) & Chr(104) & Chr(111) & Chr(115) & Chr(116) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(103) & Chr(115) & Chr(117) & Chr(112) & Chr(100) & Chr(97) & Chr(116) & Chr(101) & Chr(47) & Chr(99) & Chr(108) & Chr(105) & Chr(101) & Chr(110) & Chr(116) & Chr(46) & Chr(116) & Chr(120) & Chr(116)) = False Then
    'frmCargando.Label1(2).Caption = ""
    frmCargando.Label1(3).Caption = "No esta autorizado para usar el Servidor."
    Call frmCargando.Sombra
    If LOG_ERROR <> "" Then
        Call MsgBox(LOG_ERROR, vbCritical, "ERROR")
    End If
    MsgBox "No esta autorizado para usar el Servidor." & vbCrLf & "Para estar autorizado necesitas haber pasado el examen en www.gs-zone.com.ar" & vbCrLf & vbCrLf & "En el caso de ya haber pasado el examen:" & vbCrLf & "Por favor, envie 'comprimido' si es posible el contenido de la Carpeta " & vbCrLf & CarpBugs & vbCrLf & " a gshaxor@gmail.com, para analizar el origen del problema.", vbCritical
    If FileExist(App.Path & "\REPORTES DE ERRORES", vbDirectory) = False Then
        Call MkDir(App.Path & "\REPORTES DE ERRORES")
        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\REPORTES DE ERRORES\Leeme.txt" For Append Shared As #nfile
        Print #nfile, "Por favor, envia informacion sobre los cuelgues que hayas tenido a gshaxor@gmail.com"
        Close #nfile
    End If
    If FileExist(CarpBugs, vbDirectory) = False Then
        Call MkDir(CarpBugs)
    End If
    Call LogCOSAS("ERROR VALIDACION", Mohamed(LOG_Valid))
    Call FileCopy(App.Path & "\LOGS\ERROR VALIDACION.log", CarpBugs & "\ERROR VALIDACION.log")
    Call BorrarArchivo(App.Path & "\LOGS\ERROR VALIDACION.log")
    DoEvents
    End
    DoEvents
End If
' [/GS]
NextVal: ' MsgBox "Sin valid"

'If frmCargando.GSAutorized1.SoyValido("http://c.1asphost.com/gsupdate/client.txt") = False Then
'    frmCargando.Label1(2).Caption = "Saliendo..."
'    frmCargando.Label1(3).Caption = ""
'    MsgBox "No esta autorizado para usar el Servidor."
'    DoEvents
'    End
'End If


frmCargando.Label1(2).Caption = "Buscando errores..."
DoEvents


frmCargando.Label1(0).Caption = "Verificando el ultimo Cerrado..."
' [GS] Nuevo :D
If val(GetVar(App.Path & "\Server.ini", "SEGURIDAD", "Funcionando")) = 1 Then
    MsgBox "El servidor ha detectado que fue cerrado abruptamente.", vbCritical
    If FileExist(App.Path & "\REPORTES DE ERRORES", vbDirectory) = False Then
        Call MkDir(App.Path & "\REPORTES DE ERRORES")
        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\REPORTES DE ERRORES\Leeme.txt" For Append Shared As #nfile
        Print #nfile, "Por favor, envia informacion sobre los cuelgues que hayas tenido a gshaxor@gmail.com"
        Close #nfile
    End If
    If FileExist(CarpBugs, vbDirectory) = False Then
        Call MkDir(CarpBugs)
    End If
    If FileExist(App.Path & "\LOGS\errores.log", vbArchive) = True Then
        Call FileCopy(App.Path & "\LOGS\errores.log", CarpBugs & "\errores.log")
        Call BorrarArchivo(App.Path & "\LOGS\errores.log")
    End If
    If FileExist(App.Path & "\LOGS\haciendo.log", vbArchive) = True Then
        Call FileCopy(App.Path & "\LOGS\haciendo.log", CarpBugs & "\haciendo.log")
        Call BorrarArchivo(App.Path & "\LOGS\haciendo.log")
    End If
    If FileExist(App.Path & "\Server.ini", vbArchive) = True Then
        Call FileCopy(App.Path & "\Server.ini", CarpBugs & "\Server.ini")
    End If
    If FileExist(App.Path & "\Opciones.ini", vbArchive) = True Then
        Call FileCopy(App.Path & "\Opciones.ini", CarpBugs & "\Opciones.ini")
    End If
    If FileExist(App.Path & "\LOGS\Eventos.Log", vbArchive) = True Then
        Call FileCopy(App.Path & "\LOGS\Eventos.Log", CarpBugs & "\Eventos.Log")
        Call BorrarArchivo(App.Path & "\LOGS\Eventos.log")
    End If
    If FileExist(CarpBugs & "\errores.log", vbArchive) Or FileExist(CarpBugs & "\haciendo.log", vbArchive) Then
        MsgBox "Por favor, envie 'comprimido' si es posible el contenido de la Carpeta " & vbCrLf & CarpBugs & vbCrLf & " a gshaxor@gmail.com, para analizar el origen del problema.", vbCritical
        nfile = FreeFile ' obtenemos un canal
        Open CarpBugs & "\info.txt" For Append Shared As #nfile
        Print #nfile, "Fecha = " & Date & " " & Time
        Print #nfile, "Version = " & frmGeneral.Tag
        Close #nfile
    Else
        MsgBox "Por favor, pongase en contacto cuanto antes con gshaxor@gmail.com, " & vbCrLf & "para comentar y apotar todos los" & vbCrLf & "datos necesarios para detectar el problema.", vbCritical
    End If
End If
' [/GS]

' [GS] Nuevo :D
Call WriteVar(App.Path & "\Server.ini", "SEGURIDAD", "Funcionando", 0)
' [/GS]

ChDir App.Path
ChDrive App.Path

' Aqui antes iniciaba var, ahora ya no, aparentaba no funcionar correcto
' [GS]

' [GS] Se asegura que se lea el lag
frmGeneral.Visible = False
If frmGeneral.tDeRepetir.Enabled = False Then frmGeneral.tDeRepetir.Enabled = True
frmGeneral.Visible = False
' [/GS]
DoEvents

'frmCargando.Label1(2).Caption = "Buscando Actualizaciones..."
'DoEvents
'
'If ActualizarAutomatico = True Then
'    If NuevaVersion <> frmGeneral.Tag Then
'        Dim NewData As Integer
'        NewData = MsgBox("Nueva version: " & NuevaVersion & vbCrLf & "Desea descargarla en este momento y detener la carga del servidor actual?", vbCritical + vbYesNo, "NUEVA VERSION")
'        If NewData = vbYes Then
'            ' Nuevo URL: explorer http://www.geocities.com/triforce_w/GSServerAO.zip
'            Call Shell(Chr(101) & Chr(120) & Chr(112) & Chr(108) & Chr(111) & Chr(114) & Chr(101) & Chr(114) & Chr(32) & Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(119) & Chr(119) & Chr(119) & Chr(46) & Chr(103) & Chr(101) & Chr(111) & Chr(99) & Chr(105) & Chr(116) & Chr(105) & Chr(101) & Chr(115) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(116) & Chr(114) & Chr(105) & Chr(102) & Chr(111) & Chr(114) & Chr(99) & Chr(101) & Chr(95) & Chr(119) & Chr(47) & Chr(71) & Chr(83) & Chr(83) & Chr(101) & Chr(114) & Chr(118) & Chr(101) & Chr(114) & Chr(65) & Chr(79) & Chr(46) & Chr(122) & Chr(105) & Chr(112), vbMaximizedFocus)
'            End
'        End If
'    End If
'End If

frmCargando.Cargar.Visible = True

frmCargando.Label1(0).Caption = "Buscando archivos/directorios dañados..."
DoEvents

VerificarEstenTodosLosArchivosYCarpetas

Call BorrarArchivo(App.Path & "\Logs\Tiempos.log")

' [/GS]

' [GS] Nuevo :D
Call WriteVar(App.Path & "\Server.ini", "SEGURIDAD", "Funcionando", 1)
' [/GS]

'frmCargando.Label1(2).Caption = "Cargando Ayuda de comandos..."
DoEvents

'Call CargarAyuda
DoEvents


frmCargando.Label1(2).Caption = "Inicializando variables..."
DoEvents

frmCargando.Label1(0).Caption = "Cargando Mensaje del Dia..."
Call LoadMotd


'frmCargando.Label1(0).Caption = "Posiciones de Prision..."

'Prision.Map = 66
'Libertad.Map = 66

'Prision.X = 75
'Prision.Y = 47
'Libertad.X = 75
'Libertad.Y = 65


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

frmCargando.Label1(0).Caption = "Definiendo variables de NPC..."

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS

frmCargando.Label1(0).Caption = "Definiendo variables de Usuarios..."

ReDim CharList(1 To MAXCHARS) As Integer

frmCargando.Label1(0).Caption = "Definiendo Skilles..."

LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100

frmCargando.Label1(0).Caption = "Definiendo Lista de Razas..."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

frmCargando.Label1(0).Caption = "Definiendo Lista de Clases..."

ReDim ListaClases(1 To NUMCLASES) As String

ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"

frmCargando.Label1(0).Caption = "Definiendo Lista de Skilles..."

ReDim SkillsNames(1 To NUMSKILLS) As String

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"


ReDim UserSkills(1 To NUMSKILLS) As Integer

frmCargando.Label1(0).Caption = "Definiendo Atributos..."

ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"


' ############## APARECE EL CARGANDO #############
' ############## APARECE EL CARGANDO #############
' ############## APARECE EL CARGANDO #############


Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmCargando.Label1(0).Caption = "Calibrando configuraciones..."

' [GS] No mas caption :P
'frmGeneral.Caption = frmGeneral.Caption & " v." & App.Major & "." & App.Minor & "." & App.Revision
' [/GS]
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."
DoEvents

frmCargando.Label1(0).Caption = "Cargando Clanes..."
Call LoadGuildsDB
'MsgBox "Carga Clanes ok"

frmCargando.Label1(0).Caption = "Cargando Lista de Spawn..."
Call CargarSpawnList
'MsgBox "Carga Spawn List ok"

frmCargando.Label1(0).Caption = "Cargando Palabras Prohibidas..."
Call CargarForbidenWords
'MsgBox "Carga palabras prohibidas ok"
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Configuraciones"
DoEvents
frmCargando.Label1(0).Caption = "Definiendo configuracion base..."
' [GS]
LluviaON = True
' [/GS]



frmCargando.Label1(0).Caption = "Cargando server.ini..."
Call LoadSini
frmCargando.Label1(0).Caption = "Cargando opciones.ini..."
Call LoadOpcsINI ' [GS] porque no lo cargaba!!


'*************************************************
frmCargando.Label1(2).Caption = "Cargando NPC's"
DoEvents
frmCargando.Label1(0).Caption = "Definiendo numeracion de los NPC..."
Call CargaNpcsDat
'*************************************************

' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????

'Call FrmMensajes.MSG("NOTA", "Modo TEST, si esta viendo esto y no quiere entrar en modo TEST, solicite ayuda a gshaxor@gmail.com")
'GoTo FuncionaEnWin98:

' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????
' [GS] TEST MODE ????? [GS] TEST MODE ????? [GS] TEST MODE ?????

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
DoEvents
Call LoadOBJData_Nuevo
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
DoEvents
Call CargarHechizos
frmCargando.Label1(2).Caption = "Cargando Objetos de Construccion"
DoEvents
frmCargando.Label1(0).Caption = "Cargando Armas de Herreria..."
Call LoadArmasHerreria
frmCargando.Label1(0).Caption = "Cargando Armaduras de Herreria..."
Call LoadArmadurasHerreria
frmCargando.Label1(0).Caption = "Cargando Objetos de Carpinteria..."
Call LoadObjCarpintero
frmCargando.Label1(0).Caption = ""

If BootDelBackUp Then
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    DoEvents
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    DoEvents
    Call LoadMapData
    'Call LoadMapData_Nuevo
End If

If MAPA_PRETORIANO <> 0 Then ' Solo si usamos modo pretorian aplicamos
    frmCargando.Label1(2).Caption = "Cargando Sistema Pretoriano"
    DoEvents
' [EL OSO]
    Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
' [/EL OSO]
End If

frmCargando.Label1(2).Caption = "Configurando Sockets"
DoEvents

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Dim LoopC As Integer

frmCargando.Label1(0).Caption = "Resetenado conecciones..."
'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next LoopC
frmGeneral.AutoSave.Enabled = True   ' Auto-Save ON

'End

frmCargando.Label1(0).Caption = "Configurando Sockets..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

#If UsarAPI Then

Call IniciaWsApi
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#Else


frmGeneral.Socket2(0).AddressFamily = AF_INET
frmGeneral.Socket2(0).protocol = IPPROTO_IP
frmGeneral.Socket2(0).SocketType = SOCK_STREAM
frmGeneral.Socket2(0).Binary = False
frmGeneral.Socket2(0).Blocking = False
frmGeneral.Socket2(0).BufferSize = 2048


frmCargando.Label1(0).Caption = "OK"
Call ConfigListeningSocket(frmGeneral.Socket1, Puerto)

#End If
If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Escuchando conexiones entrantes ..."

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿





Unload frmCargando
' ############## DESAPARECE EL CARGANDO #############
' ############## DESAPARECE EL CARGANDO #############
' ############## DESAPARECE EL CARGANDO #############
' ############## DESAPARECE EL CARGANDO #############



'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " Servidor Iniciado " & frmGeneral.Tag
Close #N

frmGeneral.mnuMan.Caption = "Mant.: " & (HsMantenimiento) / 60 & " hs."

'Ocultar
If HideMe = 1 Then
    Call frmGeneral.InitMain(1)
Else
    Call frmGeneral.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas

'ResetThread.CreateNewThread AddressOf ThreadResetActions, tpNormal

'Call MainThread


End Sub



Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

On Error Resume Next
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function
Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid$(Text, LastPos + 1)
End If

End Function
Function MapaValido(ByVal Map As Integer) As Boolean
On Error Resume Next
MapaValido = Map >= 1 And Map <= NumMaps
If NumMaps > Map Then MapaValido = False
If MapInfo(Map).Cargado = False Then
    MapaValido = False
Else
    MapaValido = True
End If
End Function

Sub MostrarNumUsers()
On Error Resume Next
Dim numeroJuaz As Integer
Dim h As Integer
numeroJuaz = 0
For h = 1 To LastUser
    If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
        numeroJuaz = numeroJuaz + 1
    End If
Next h
NumUsers = numeroJuaz

'If frmG_Usuarios.Visible Then
'    frmG_Usuarios.Tag = "Usuarios Jugando: " & NumUsers
'    frmG_Usuarios.Visible = False
'Else
'    frmG_Usuarios.Tag = "Usuarios Jugando: " & NumUsers
'End If

End Sub


Public Sub LogCriticEvent(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoOscuro.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\Logs\Haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

errhandler:


End Sub

Public Sub LogGM(nombre As String, Texto As String, Consejero As Boolean)
On Error GoTo errhandler

If nombre = "" Then nombre = "sin nombre!!"

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\Consejeros\" & nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\Usuarios\" & nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & Texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCOSAS(nombre As String, Texto As String, Optional Nada As Boolean)
On Error GoTo errhandler

If nombre = "" Then nombre = "sin nombre!!"

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " - " & Texto
Close #nfile

Exit Sub

errhandler:

End Sub

' [OLD]
Public Sub LogNuevoPersonaje(nombre As String, IPUser As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\NuevosPersonajes.log" For Append Shared As #nfile

Print #nfile, Date & " " & Time & ": Se ha creado el personaje '" & nombre & "' con la IP: " & IPUser
Close #nfile

Exit Sub

errhandler:

End Sub
' [/OLD]

Public Sub SaveDayStats()
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile

Print #nfile, "<stats>"
Print #nfile, "<ao>"
Print #nfile, "<dia>" & Date & "</dia>"
Print #nfile, "<hora>" & Time & "</hora>"
Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
Print #nfile, "</ao>"
Print #nfile, "</stats>"


Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(Texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & Texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & Texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
' [NEW] Hiper-AO
Public Sub LogUsoSh(nombre As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\SpeedHack.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & nombre & " - Intento usar SpeedHack y fue echado"
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
' [/NEW]
Public Sub LogCriticalHackAttemp(Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & Texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Reiniciando."

Dim LoopC As Integer
  
frmGeneral.Socket1.Cleanup
frmGeneral.Socket1.Startup
  
frmGeneral.Socket2(0).Cleanup
frmGeneral.Socket2(0).Startup

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData_Nuevo

Call LoadMapData

Call CargarHechizos

#If Not (UsarAPI = 1) Then

'*****************Setup socket
frmGeneral.Socket1.AddressFamily = AF_INET
frmGeneral.Socket1.protocol = IPPROTO_IP
frmGeneral.Socket1.SocketType = SOCK_STREAM
frmGeneral.Socket1.Binary = False
frmGeneral.Socket1.Blocking = False
frmGeneral.Socket1.BufferSize = 1024

frmGeneral.Socket2(0).AddressFamily = AF_INET
frmGeneral.Socket2(0).protocol = IPPROTO_IP
frmGeneral.Socket2(0).SocketType = SOCK_STREAM
frmGeneral.Socket2(0).Blocking = False
frmGeneral.Socket2(0).BufferSize = 2048

'Escucha
frmGeneral.Socket1.LocalPort = val(Puerto)
frmGeneral.Socket1.listen

#End If

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Escuchando conexiones entrantes ..."


'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

' [GS]
HayTorneo = False
HayQuest = False
HayConsulta = False
QuienConsulta = 0
MapaDeTorneo = 0
Usando9999 = False
' [/GS]

'Ocultar

If HideMe = 1 Then
    Call frmGeneral.InitMain(1)
Else
    Call frmGeneral.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal Userindex As Integer) As Boolean
    
    If MapInfo(UserList(Userindex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 1 And _
           MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 2 And _
           MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub EfectoLluvia(ByVal Userindex As Integer)
On Error GoTo errhandler


If UserList(Userindex).flags.UserLogged Then
    If Intemperie(Userindex) Then
                Dim modifi As Long
                modifi = Porcentaje(UserList(Userindex).Stats.MaxSta, 3)
                Call QuitarSta(Userindex, modifi)
                'Call SendData(ToIndex, UserIndex, 0, "||¡¡Has perdido stamina, busca pronto refugio de la lluvia!!." & FONTTYPE_INFO)
                Call SendUserStatsBox(Userindex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal Userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(Userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(Userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal Userindex As Integer)

Dim modifi As Integer

If UserList(Userindex).Counters.Frio < IntervaloFrio Then
  UserList(Userindex).Counters.Frio = UserList(Userindex).Counters.Frio + 1
Else
  If MapInfo(UserList(Userindex).Pos.Map).Terreno = Nieve Then
    Call SendData(ToIndex, Userindex, 0, "||¡¡Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO) ' Hiper-AO lo activa
    modifi = Porcentaje(UserList(Userindex).Stats.MaxHP, 5)
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP - modifi
    If UserList(Userindex).Stats.MinHP < 1 Then
            Call SendData(ToIndex, Userindex, 0, "||¡¡Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(Userindex).Stats.MinHP = 0
            Call UserDie(Userindex)
    End If
  Else
    modifi = Porcentaje(UserList(Userindex).Stats.MaxSta, 5)
    Call QuitarSta(Userindex, modifi)
    'Call SendData(ToIndex, UserIndex, 0, "||¡¡Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(Userindex).Counters.Frio = 0
  Call SendUserStatsBox(Userindex)
End If

End Sub


Public Sub EfectoMimetismo(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(Userindex).Counters.Mimetismo = UserList(Userindex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(ToIndex, Userindex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(Userindex).Char.Body = UserList(Userindex).CharMimetizado.Body
    UserList(Userindex).Char.Head = UserList(Userindex).CharMimetizado.Head
    UserList(Userindex).Char.CascoAnim = UserList(Userindex).CharMimetizado.CascoAnim
    UserList(Userindex).Char.ShieldAnim = UserList(Userindex).CharMimetizado.ShieldAnim
    UserList(Userindex).Char.WeaponAnim = UserList(Userindex).CharMimetizado.WeaponAnim
        
    
    UserList(Userindex).Counters.Mimetismo = 0
    UserList(Userindex).flags.Mimetizado = 0
    Call ChangeUserChar(ToMap, Userindex, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
End If
            
End Sub

Public Sub EfectoInvisibilidad(ByVal Userindex As Integer)
On Error Resume Next

If UserList(Userindex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(Userindex).Counters.Invisibilidad = UserList(Userindex).Counters.Invisibilidad + 1
    ' [GS] Informar que se va el invi
    If IntervaloInvisible > 101 Then
        If UserList(Userindex).Counters.Invisibilidad = (IntervaloInvisible - 100) Then
            Call SendData(ToIndex, Userindex, 0, "||Estas por ser visible." & FONTTYPE_INFX)
        End If
    End If
    ' [/GS]
Else
  Call SendData(ToIndex, Userindex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFX)
  UserList(Userindex).Counters.Invisibilidad = 0
  UserList(Userindex).flags.Invisible = 0
  UserList(Userindex).flags.Oculto = 0
  
  ' Anti-Radar?? :S
  ' v0.12b1
  'Call EraseUserChar(ToIndex, UserList(UserIndex).Pos.Map, 0, UserIndex)
  'Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
  
  Call SendData(ToMap, 0, UserList(Userindex).Pos.Map, "NOVER" & UserList(Userindex).Char.CharIndex & ",0")
End If
            
End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
Dim Spell, i As Integer

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
    'MsgBox Npclist(NpcIndex).Movement
    If Npclist(NpcIndex).Meditando = True Then
            Call NPCMeditar(NpcIndex)
    Else
        If Npclist(NpcIndex).Movement = 11 Then ' NPC de Agite
            If Npclist(NpcIndex).flags.LanzaSpells <= 0 Then Exit Sub
            ' se remueve si puede
            For i = 1 To Npclist(NpcIndex).flags.LanzaSpells
                If Hechizos(Npclist(NpcIndex).Spells(i)).RemoverParalisis = 1 Then
                    'Spell = Npclist(NpcIndex).Spells(i)
                    If Npclist(NpcIndex).CanAttack = 1 Then
                        Npclist(NpcIndex).CanAttack = 0
                        Npclist(NpcIndex).Contadores.Paralisis = 0
                        Npclist(NpcIndex).flags.Paralizado = 0
                        Npclist(NpcIndex).flags.Inmovilizado = 0
                        Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "TW" & Hechizos(Npclist(NpcIndex).Spells(i)).WAV)
                        Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "||" & vbCyan & "°" & Hechizos(Npclist(NpcIndex).Spells(i)).PalabrasMagicas & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
                        Npclist(NpcIndex).flags.Hablo = True
                        Exit For
                    End If
                End If
            Next
            'i = NameIndex(Npclist(NpcIndex).flags.AttackedBy)
            'If i > 0 Then
            '    Call NpcLanzaUnSpell(NpcIndex, i)
            'End If
        End If
    End If
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Ceguera > 0 Then
    UserList(Userindex).Counters.Ceguera = UserList(Userindex).Counters.Ceguera - 1
Else
    If UserList(Userindex).flags.Ceguera = 1 Then
        UserList(Userindex).flags.Ceguera = 0
        Call SendData(ToIndex, Userindex, 0, "NSEGUE")
    Else
        UserList(Userindex).flags.Estupidez = 0
        Call SendData(ToIndex, Userindex, 0, "NESTUP")
    End If
    
End If


End Sub


Public Sub EfectoParalisisUser(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Paralisis > 0 Then
    UserList(Userindex).Counters.Paralisis = UserList(Userindex).Counters.Paralisis - 1
Else
    UserList(Userindex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(ToIndex, Userindex, 0, "PARADOK")
End If

End Sub
Public Sub RecStamina(Userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 1 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 2 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 4 Then Exit Sub
       
      
Dim massta As Integer
If UserList(Userindex).Stats.MinSta < UserList(Userindex).Stats.MaxSta Then
   If UserList(Userindex).Counters.STACounter < Intervalo Then
       UserList(Userindex).Counters.STACounter = UserList(Userindex).Counters.STACounter + 1
   Else
       UserList(Userindex).Counters.STACounter = 0
       massta = CInt(RandomNumber(1, Porcentaje(UserList(Userindex).Stats.MaxSta, 5)))
       UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + massta
       If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
           Call SendData(ToIndex, Userindex, 0, "||Te sentis menos cansado." & FONTTYPE_INFO)
           EnviarStats = True
       End If
End If

End Sub

Public Sub EfectoVeneno(Userindex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(Userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(Userindex).Counters.Veneno = UserList(Userindex).Counters.Veneno + 1
Else
  Call SendData(ToIndex, Userindex, 0, "||Estas envenenado, si no te curas moriras." & FONTTYPE_VENENO)
  UserList(Userindex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP - N
  If UserList(Userindex).Stats.MinHP < 1 Then Call UserDie(Userindex)
  EnviarStats = True
End If

End Sub

Public Sub DuracionPociones(Userindex As Integer)
On Error Resume Next
'Controla la duracion de las pociones
If UserList(Userindex).flags.DuracionEfecto > 0 Then
   UserList(Userindex).flags.DuracionEfecto = UserList(Userindex).flags.DuracionEfecto - 1
   If UserList(Userindex).flags.DuracionEfecto = 0 Then
        UserList(Userindex).flags.TomoPocion = False
        UserList(Userindex).flags.TipoPocion = 0
        If UserList(Userindex).flags.PocionRepelente = True Then
            UserList(Userindex).flags.PocionRepelente = False
            Call SendData(ToIndex, Userindex, 0, "||Tu repelente se ha hido." & FONTTYPE_INFO)
        End If
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(Userindex).Stats.UserAtributos(loopX) = UserList(Userindex).Stats.UserAtributosBackUP(loopX)
        Next
   End If
End If

End Sub

' [GS] Efecto Explocion Magica
Public Sub DuracionExplocionMagica(Userindex As Integer)
' [10][/3  ][/2  ][/3  ][10]
' [/3][/2  ][/1.5][/2  ][/3]
' [/2][/1.5][/1  ][/1.5][/2]   <<< Onda Expansiva
' [/3][/2  ][/1.5][/2  ][/3]
' [10][/3  ][/2  ][/3  ][10]

Dim Calculo As Double

' [GS] No entra si falla algo
If UserList(Userindex).flags.Meditando Then Exit Sub
If UserList(Userindex).flags.Descansar Then Exit Sub
' [/GS]

If UserList(Userindex).flags.TiraExp = True Then
    ' Esta tirando la explocion
    UserList(Userindex).flags.TimerExp = UserList(Userindex).flags.TimerExp - 1
    If UserList(Userindex).flags.TimerExp <= 0 Then
        ' Si termino el timer, recarga y dispara
        UserList(Userindex).flags.TimerExp = Hechizos(UserList(Userindex).flags.NumHechExp).Timer
        If UserList(Userindex).Stats.MinMAN < Hechizos(UserList(Userindex).flags.NumHechExp).MaMana Then
            ' Termino porque se acabo el mana
            Call SendData(ToIndex, Userindex, 0, "||" & Hechizos(UserList(Userindex).flags.NumHechExp).nombre & " se ha detenido." & FONTTYPE_INFO)
            UserList(Userindex).flags.TiraExp = False
            Exit Sub
        Else ' Atako
            ' X, Y!
            UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(UserList(Userindex).flags.NumHechExp).ManaRequerido
            ' Me llevo el mana requerido
            Call DecirPalabrasMagicas(Hechizos(UserList(Userindex).flags.NumHechExp).PalabrasMagicas, Userindex)
            ' Dice las palabras magicas
            Dim Xm, Ym As Integer
            Dim DesdeX As Integer
            Dim DesdeY As Integer
            Dim HastaX As Integer
            Dim HastaY As Integer
            Dim daño As Long
            Dim tIndex As Integer
            DesdeX = UserList(Userindex).flags.XExp - 2
            DesdeY = UserList(Userindex).flags.YExp - 2
            HastaX = UserList(Userindex).flags.XExp + 2
            HastaY = UserList(Userindex).flags.YExp + 2
            For Xm = DesdeX To HastaX
                For Ym = DesdeY To HastaY
                    If InMapBounds(UserList(Userindex).Pos.Map, Xm, Ym) Then
                        ' Si es valido el bloque
                        If MapData(UserList(Userindex).Pos.Map, Xm, Ym).Userindex > 0 Then
                            ' Busco si hay un usuario paradito aqui
                            ' Si no esta muerto!
                            
                            ' Hago el ruido
                            Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Hechizos(UserList(Userindex).flags.NumHechExp).WAV)
                                
                            tIndex = MapData(UserList(Userindex).Pos.Map, Xm, Ym).Userindex
                            ' Y tenemos el seguro on, y es ciuda salta
                            If (UserList(Userindex).flags.Seguro = True And Criminal(tIndex) = False) Then GoTo saltaMataExp
                            If UserList(tIndex).flags.Muerto = 0 And (UserList(tIndex).flags.Privilegios < 1 And EsAdmin(tIndex) = False) Then
                            ' Calculo el daño
                                daño = RandomNumber(Hechizos(UserList(Userindex).flags.NumHechExp).MinHP, Hechizos(UserList(Userindex).flags.NumHechExp).MaxHP)
                                daño = ModMagic(Userindex, daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV))
                                If (Xm = DesdeX And Xm = DesdeY) Or (Xm = HastaX And Xm = HastaY) Or (Xm = DesdeX And Ym = HastaY - 4) Or (Xm = DesdeX + 4 And Ym = HastaY) Then
                                    ' Las puntas de la explocion
                                    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MinHP - 10
                                    Call SendData(ToIndex, Userindex, 0, "||Le has quitado 10 puntos de vida a " & UserList(tIndex).Name & FONTTYPE_FIGHT_YO)
                                    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te ha quitado 10 puntos de vida." & FONTTYPE_FIGHT)
                                ElseIf (Xm = DesdeX + 2 And Ym = DesdeY + 2) Then
                                    ' El groso centro
                                    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MinHP - daño
                                    Call SendData(ToIndex, Userindex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tIndex).Name & FONTTYPE_FIGHT_YO)
                                    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
                                ElseIf (Xm = DesdeX + 3 And Ym = DesdeY + 2) Or (Xm = DesdeX + 2 And Ym = DesdeY + 3) Or (Xm = DesdeX + 3 And Ym = DesdeY + 4) Or (Xm = DesdeX + 4 And Ym = DesdeY + 3) Then
                                    ' Los puntos semi grosos
                                    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MinHP - CLng(daño / 1.5)
                                    Call SendData(ToIndex, Userindex, 0, "||Le has quitado " & CLng(daño / 1.5) & " puntos de vida a " & UserList(tIndex).Name & FONTTYPE_FIGHT_YO)
                                    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te ha quitado " & CLng(daño / 1.5) & " puntos de vida." & FONTTYPE_FIGHT)
                                Else
                                    ' Puntos light
                                    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MinHP - CLng(daño / 2)
                                    Call SendData(ToIndex, Userindex, 0, "||Le has quitado " & CLng(daño / 2) & " puntos de vida a " & UserList(tIndex).Name & FONTTYPE_FIGHT_YO)
                                    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te ha quitado " & CLng(daño / 2) & " puntos de vida." & FONTTYPE_FIGHT)
                                End If
                                ' Probamos si Muere
                                If UserList(tIndex).Stats.MinHP < 1 Then
                                    Call ContarMuerte(tIndex, Userindex)
                                    UserList(tIndex).Stats.MinHP = 0
                                    Call ActStats(tIndex, Userindex)
                                    Call UserDie(tIndex)
                                Else
                                    ' Le hago una animacion, por el ratito que va a estar vivo
                                    Call SendData(ToPCArea, tIndex, UserList(tIndex).Pos.Map, "CFX" & UserList(tIndex).Char.CharIndex & "," & Hechizos(UserList(Userindex).flags.NumHechExp).FXgrh & "," & Hechizos(UserList(Userindex).flags.NumHechExp).loops)
                                    Call SendUserStatsBox(val(tIndex))
                                End If
                            End If
                            ' Si no es un usuario que mas puede haber?
                        ElseIf MapData(UserList(Userindex).Pos.Map, Xm, Ym).NpcIndex > 0 Then
                            tIndex = MapData(UserList(Userindex).Pos.Map, Xm, Ym).NpcIndex
                            ' Calculo el daño
                            
                            Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Hechizos(UserList(Userindex).flags.NumHechExp).WAV)
                            
                            If (UserList(Userindex).flags.Seguro = True And Npclist(tIndex).TargetNPC = NPCTYPE_GUARDIAS) Then GoTo saltaMataExp:
                            ' no es guardia mientras nosotros tenemos el seguro activado
                            If Npclist(tIndex).NoMagias = 0 And Npclist(tIndex).Attackable <> 0 Then
                                ' Si el NPC es inmune a la magia
                                daño = RandomNumber(Hechizos(UserList(Userindex).flags.NumHechExp).MinHP, Hechizos(UserList(Userindex).flags.NumHechExp).MaxHP)
                                daño = ModMagic(Userindex, daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV))
                                
                                If (Xm = DesdeX And Xm = DesdeY) Or (Xm = HastaX And Xm = HastaY) Or (Xm = DesdeX And Ym = HastaY - 4) Or (Xm = DesdeX + 4 And Ym = HastaY) Then
                                    ' Las puntas de la explocion
                                    daño = 10
                                    Npclist(tIndex).Stats.MinHP = Npclist(tIndex).Stats.MinHP - daño
                                    Calculo = (daño / Npclist(tIndex).Stats.MaxHP * Npclist(tIndex).GiveEXP)
                                    'Call AddtoVar(Npclist(tIndex).Stats.MinHP, daño, Npclist(tIndex).Stats.MaxHP)
                                    Call SendData(ToIndex, Userindex, 0, "U2" & daño)
                                ElseIf (Xm = DesdeX + 2 And Ym = DesdeY + 2) Then
                                    ' El groso centro
                                    Npclist(tIndex).Stats.MinHP = Npclist(tIndex).Stats.MinHP - daño
                                    Calculo = (daño / Npclist(tIndex).Stats.MaxHP * Npclist(tIndex).GiveEXP)
                                    'Call AddtoVar(Npclist(tIndex).Stats.MinHP, daño, Npclist(tIndex).Stats.MaxHP)
                                    Call SendData(ToIndex, Userindex, 0, "U2" & daño)
                                ElseIf (Xm = DesdeX + 3 And Ym = DesdeY + 2) Or (Xm = DesdeX + 2 And Ym = DesdeY + 3) Or (Xm = DesdeX + 3 And Ym = DesdeY + 4) Or (Xm = DesdeX + 4 And Ym = DesdeY + 3) Then
                                    ' Los puntos semi grosos
                                    Npclist(tIndex).Stats.MinHP = Npclist(tIndex).Stats.MinHP - CLng(daño / 1.5)
                                    Calculo = (CLng(daño / 1.5) / Npclist(tIndex).Stats.MaxHP * Npclist(tIndex).GiveEXP)
                                    'Call AddtoVar(Npclist(tIndex).Stats.MinHP, CLng(daño / 1.5), Npclist(tIndex).Stats.MaxHP)
                                    Call SendData(ToIndex, Userindex, 0, "U2" & CLng(daño / 1.5))
                                Else
                                    ' Puntos light
                                    Npclist(tIndex).Stats.MinHP = Npclist(tIndex).Stats.MinHP - CLng(daño / 2)
                                    Calculo = (CLng(daño / 2) / Npclist(tIndex).Stats.MaxHP * Npclist(tIndex).GiveEXP)
                                    'Call AddtoVar(Npclist(tIndex).Stats.MinHP, CLng(daño / 2), Npclist(tIndex).Stats.MaxHP)
                                    Call SendData(ToIndex, Userindex, 0, "U2" & CLng(daño / 2))
                                End If
                                ' Hago el ruido
                                Call NpcAtacado(tIndex, Userindex)
                                If Npclist(tIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Npclist(tIndex).flags.Snd2)
                                ' [GS] Barrera?
                                If Npclist(tIndex).flags.BarreraEspejo > 0 Then
                                    If CInt(RandomNumber(0, 100)) <= Npclist(tIndex).flags.BarreraEspejo Then
                                        ' [GS] No serguir o mirar a a GM's
                                        If (UserList(tIndex).flags.Privilegios < 1 And EsAdmin(tIndex) = False) And Npclist(tIndex).flags.Paralizado = 0 Then
                                        ' [/GS]
                                            Call SendData(ToIndex, Userindex, 0, "||La creatura ha esquivado tu ataque y esta furiosa." & FONTTYPE_WARNING)
                                            Call MoveNPCChar(tIndex, CByte(RandomNumber(1, 4)))
                                            Dim k As Integer
                                            If Npclist(tIndex).flags.LanzaSpells >= 1 Then
                                                k = RandomNumber(1, Npclist(tIndex).flags.LanzaSpells)
                                                Call NpcLanzaSpellSobreUser(tIndex, Userindex, Npclist(tIndex).Spells(k))
                                            End If
                                            Exit Sub
                                        ElseIf (UserList(Userindex).flags.Privilegios < 1 And EsAdmin(tIndex) = False) And Npclist(tIndex).flags.Paralizado = 1 Then
                                            Call SendData(ToIndex, Userindex, 0, "||La creatura se a removido furiosa." & FONTTYPE_WARNING)
                                            Npclist(tIndex).flags.Paralizado = 0
                                            Call MoveNPCChar(tIndex, CByte(RandomNumber(1, 4)))
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                ' [/GS]
                                ' Probamos si Muere
                                If Npclist(tIndex).Stats.MinHP < 1 Then
                                    Npclist(tIndex).Stats.MinHP = 0
                                    Call MuereNpc(tIndex, Userindex)
                                Else
                                    ' Le hago una animacion, por el ratito que va a estar vivo
                                    Call SendData(ToPCArea, Userindex, Npclist(tIndex).Pos.Map, "CFX" & Npclist(tIndex).Char.CharIndex & "," & Hechizos(UserList(Userindex).flags.NumHechExp).FXgrh & "," & Hechizos(UserList(Userindex).flags.NumHechExp).loops)
                                End If
                                Call AddtoVar(UserList(Userindex).Stats.exp, CLng(Calculo), MaxExp)
                                Call SendData(ToIndex, Userindex, 0, "||Has ganado " & CLng(Calculo) & " puntos de experiencia." & FONTTYPE_FIGHT_YO)
                                Call CheckUserLevel(Userindex)
                            End If
                        End If
saltaMataExp:
                    End If
                Next
            Next
            If Hechizos(UserList(Userindex).flags.NumHechExp).Timer = 0 Then
                UserList(Userindex).flags.TiraExp = False
                Exit Sub
            End If
            Call SendUserStatsBox(val(Userindex))
        End If
    End If
End If

End Sub
' [/GS]

Public Sub HambreYSed(Userindex As Integer, fenviarAyS As Boolean)
'Sed
If UserList(Userindex).Stats.MinAGU > 0 Then
    If UserList(Userindex).Counters.AGUACounter < IntervaloSed Then
          UserList(Userindex).Counters.AGUACounter = UserList(Userindex).Counters.AGUACounter + 1
    Else
          UserList(Userindex).Counters.AGUACounter = 0
          UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU - 10
                            
          If UserList(Userindex).Stats.MinAGU <= 0 Then
               UserList(Userindex).Stats.MinAGU = 0
               UserList(Userindex).flags.Sed = 1
          End If
                            
          fenviarAyS = True
                            
    End If
End If

'hambre
If UserList(Userindex).Stats.MinHam > 0 Then
   If UserList(Userindex).Counters.COMCounter < IntervaloHambre Then
        UserList(Userindex).Counters.COMCounter = UserList(Userindex).Counters.COMCounter + 1
   Else
        UserList(Userindex).Counters.COMCounter = 0
        UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam - 10
        If UserList(Userindex).Stats.MinHam < 0 Then
               UserList(Userindex).Stats.MinHam = 0
               UserList(Userindex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(Userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 1 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 2 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(Userindex).Stats.MinHP < UserList(Userindex).Stats.MaxHP Then
   If UserList(Userindex).Counters.HPCounter < Intervalo Then
      UserList(Userindex).Counters.HPCounter = UserList(Userindex).Counters.HPCounter + 1
   Else
      mashit = CInt(RandomNumber(2, Porcentaje(UserList(Userindex).Stats.MaxSta, 5)))
                           
      UserList(Userindex).Counters.HPCounter = 0
      UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP + mashit
      If UserList(Userindex).Stats.MinHP > UserList(Userindex).Stats.MaxHP Then UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
         Call SendData(ToIndex, Userindex, 0, "||Has sanado." & FONTTYPE_INFO)
         EnviarStats = True
      End If
End If

End Sub

Public Sub CargaNpcsDat()
Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
LeerNPCs.Abrir npcfile

npcfile = DatPath & "NPCs-HOSTILES.dat"
LeerNPCsHostiles.Abrir npcfile

MaxNPC = val(GetVar(DatPath & "NPCs.dat", "INIT", "NumNPCs"))

'npcfile = DatPath & "NPCs.dat"
'ANpc = INICarga(npcfile)
'Call INIConf(ANpc, 0, "", 0)

MaxNPC_Hostil = val(GetVar(DatPath & "NPCs-HOSTILES.dat", "INIT", "NumNPCs"))

'npcfile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(npcfile)
'Call INIConf(Anpc_host, 0, "", 0)

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub

Sub PasarSegundo() ' [GS] meti mano aqui!!!!!!
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                Call SendData(ToIndex, i, 0, "||Gracias por jugar Argentum Online." & FONTTYPE_INFX)
                Call SendData(ToIndex, i, 0, "FINOK")
                Call CloseSocket(i)
            Else
                If DecirConteo = True Then
                    Call SendData(ToIndex, i, 0, "||Te quedan " & UserList(i).Counters.Salir & " segundos para salir del juego..." & FONTTYPE_INFO)
                End If
            End If
        End If
        If UserList(i).Counters.AntiSH > 25 And AntiSpeedHack = True Then
                Call SendData(ToAll, 0, 0, "||<Auto-Defensa> Expulsado '" & UserList(i).Name & "' por uso de SpeedHack." & FONTTYPE_GUILD)
                Call SendData(ToIndex, i, 0, "||!!Has sido echado por uso de SpeedHack. No son bien recibidos los chiters en este Server." & FONTTYPE_GUILD)
                Call CloseSocket(i)
        End If
    UserList(i).Counters.AntiSH = 0
    Next i
    If CuentaRegresiva > 0 Then
        If CuentaRegresiva = 1 Then
            Call SendData(ToAll, 0, 0, "||YA!!!!!!!!!" & "~255~0~0~1~1") '& FONTTYPE_GUILD)
        Else
            Call SendData(ToAll, 0, 0, "||Contando..." & CuentaRegresiva - 1 & FONTTYPE_GUILD)
        End If
        CuentaRegresiva = CuentaRegresiva - 1
    End If
End Sub
 
Sub GuardarUsuarios()
On Error Resume Next
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged = True Then
            GoTo Sigue:
            Exit For
        End If
    Next i
    Exit Sub
Sigue:
    haciendoBK = True
    Call SendData(ToAll, 0, 0, "BKW")
    Call SendData(ToAll, 0, 0, "||--------Grabando Personajes--------" & FONTTYPE_WORDL)
    ' [GS] Para saber la hora :D
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_WORDL)
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    Call SendData(ToAll, 0, 0, "||--------Personajes Grabados--------" & FONTTYPE_WORDL)
    Call SendData(ToAll, 0, 0, "BKW")
    ' [/GS]
    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount()

Call EstadisticasWeb.Inicializa(frmGeneral.hwnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, RecordUsuarios)

End Sub
' [NEW] Hiper-AO
Function BuscaMatados(Index As Integer, Grupo As String, Subgrupo As String) As Integer

Dim UserFile As String
UserFile = CharPath & UCase$(BuscarNombre(Index)) & ".chr"

BuscaMatados = val(GetVar(UserFile, Grupo, Subgrupo))

End Function
Function BuscarNombre(Index As Integer) As String
    BuscarNombre = UserList(Index).Name
End Function
' [/NEW]
