Attribute VB_Name = "ES"
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
Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpfilename As String) As Long
' Ayuda al BorrarArchivo, para borrar archivos de solo lectura

Public Function BorrarArchivo(Arch As String) As Long
On Error Resume Next
BorrarArchivo = DeleteFile(Arch)   ' 0 si falla <> 0 si esta bien borrado
End Function

Public Sub CargarSpawnList()
On Error Resume Next
    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC


End Sub

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String

If EsAdmin(NameIndex(Name)) Then Exit Function

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
' [GS]
If UCase$(Name) = "GS" Then EsDios = True
' [/GS]
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False

End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False
End Function

Function EsAyudante(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Ayudantes"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Ayudantes", "Ayudante" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsAyudante = True
        Exit Function
    End If
Next WizNum
EsAyudante = False
End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Public Function TxtDimension(ByVal Name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open Name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub
Public Sub CargarHechizos()
On Error GoTo errhandler

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsLeerInis

frmCargando.Label1(0).Caption = "Preparando..."

Leer.Abrir DatPath & "Hechizos.dat"
'j = Val(Leer.DarValor(

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.DarValor("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.Cargar.Min = 0
frmCargando.Cargar.max = NumeroHechizos
frmCargando.Cargar.Value = 0
frmGeneral.ProG1.Min = 0
frmGeneral.ProG1.max = NumeroHechizos
frmGeneral.ProG1.Value = 0
frmGeneral.ProG1.Visible = True
'Llena la lista
For Hechizo = 1 To NumeroHechizos
    If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando Hechizos. " & str(Hechizo) & " de " & NumeroHechizos
    Hechizos(Hechizo).nombre = Leer.DarValor("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).desc = Leer.DarValor("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.DarValor("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.DarValor("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.DarValor("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.DarValor("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Target = val(Leer.DarValor("Hechizo" & Hechizo, "Target"))
    Hechizos(Hechizo).Tipo = val(Leer.DarValor("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.DarValor("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.DarValor("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.DarValor("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(Leer.DarValor("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.DarValor("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.DarValor("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.DarValor("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.DarValor("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.DarValor("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.DarValor("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.DarValor("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.DarValor("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.DarValor("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverParalisis"))
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.DarValor("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.DarValor("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.DarValor("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.DarValor("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.DarValor("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.DarValor("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.DarValor("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.DarValor("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.DarValor("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(Leer.DarValor("Hechizo" & Hechizo, "Cant"))
    
    
    Hechizos(Hechizo).Materializa = val(Leer.DarValor("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).ItemIndex = val(Leer.DarValor("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.DarValor("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "ManaRequerido"))

    ' [GS] Tipo invalido
    If Hechizos(Hechizo).Tipo < 1 Or Hechizos(Hechizo).Tipo > 5 Then
        If Hechizos(Hechizo).nombre <> "" Then Call Alerta("El Hechizo " & Hechizo & " tiene un Tipo invalido. Tipo: " & Hechizos(Hechizo).Tipo)
    End If
    ' [/GS]
    
    ' [GS] Nuevos tipos de magia
    Hechizos(Hechizo).RemoverCeguera = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverCeguera"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverEstupidez"))
    ' [GS]
    
    ' [GS]
    Hechizos(Hechizo).ExclusivoClase = Clase2Num(Leer.DarValor("Hechizo" & Hechizo, "ExclusivoClase"))
    ' [/GS]
    
    ' [GS] Timer Magia Explosiva
    Hechizos(Hechizo).Timer = val(Leer.DarValor("Hechizo" & Hechizo, "Timer"))
    If Hechizos(Hechizo).Timer < 0 Then Hechizos(Hechizo).Timer = 0
    ' [/GS]
        
    ' [GS] Target invalido
    If Hechizos(Hechizo).Target < 0 Or Hechizos(Hechizo).Target > 4 Then
        Call Alerta("El Hechizo " & Hechizo & " tiene un Target invalido.")
    End If
    ' [/GS]
    
    ' [GS] Requiere
    Hechizos(Hechizo).Requiere = val(Leer.DarValor("Hechizo" & Hechizo, "Requiere"))
    If Hechizos(Hechizo).Requiere > 0 Then
        If Hechizos(Hechizo).Requiere > NumObjDatas Then
            Call Alerta("El Hechizo " & Hechizo & " requiere el numero de objeto " & Hechizos(Hechizo).Requiere & " y no existe.")
        End If
    Else
        Hechizos(Hechizo).Requiere = 0
    End If
    ' [/GS]
    
    ' v0.12a9
    
    Hechizos(Hechizo).Mimetiza = val(Leer.DarValor("hechizo" & Hechizo, "Mimetiza"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.DarValor("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.DarValor("Hechizo" & Hechizo, "Inmoviliza"))
    
    Hechizos(Hechizo).StaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).NeedStaff = val(Leer.DarValor("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.DarValor("Hechizo" & Hechizo, "StaffAffected")))
    
    If frmCargando.Visible Then
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
        frmCargando.Label1(0).Caption = Hechizos(Hechizo).nombre
    End If
    If frmGeneral.Visible Then frmGeneral.ProG1.Value = frmGeneral.ProG1.Value + 1
Next
frmGeneral.ProG1.Visible = False

'Hacer Copia de de Seguridad
Call HacerDAT("Hechizos.dat")

Exit Sub

errhandler:

Call LogError("Error en CargarHechizos: " & Err.Number & " " & Err.Description)
Call RepararDAT("Hechizos.dat", "Hechizos_2.dat")

End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines")) + 1
ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).Texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = ""
Next i



'MOTD(MaxLines).texto = "**Bienvenido a " & GetVar(App.Path & "\server.ini", "INIT", "ServerName") & " ** - Servidor " & frmMain.Label2.Tag & " programado por ^[GS]^, basado en ZipAO (Sicarul)~255~255~0~1~0"
'MOTD(MaxLines).Formato = ""
End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")

If haciendoBK = True Then Exit Sub

frmGeneral.mnuArchivo.Enabled = False
frmGeneral.mnuVer.Enabled = False
frmGeneral.mnuActualizar.Enabled = False
frmGeneral.mnuAcciones.Enabled = False
frmGeneral.mnuPopupMenu.Enabled = False

haciendoBK = True

Call SendData(ToAll, 0, 0, "||INCIANDO WORLDSAVE, POR FAVOR ESPERE" & FONTTYPE_WORDL)
Call SendData(ToAll, 0, 0, "BKW")
DoEvents
Call SendData(ToAll, 0, 0, "||--- Guardando Clanes..." & "~32~51~223~0~0")
Call SaveGuildsDB
Call SendData(ToAll, 0, 0, "||OK" & "~223~51~32~1~1")
Call LimpiarMundo
Call WorldSave

Call SendData(ToAll, 0, 0, "BKW")

Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

If frmGeneral.mnuCerrarCorrectamente.Checked = False Then
    frmGeneral.mnuArchivo.Enabled = True
    frmGeneral.mnuVer.Enabled = True
    frmGeneral.mnuActualizar.Enabled = True
    frmGeneral.mnuAcciones.Enabled = True
    frmGeneral.mnuPopupMenu.Enabled = True
End If
'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub


Public Sub SaveMapData(ByVal N As Integer)

On Error Resume Next
'Call LogTarea("Sub SaveMapData N:" & n)

Dim LoopC As Integer
Dim TempINT As Integer
Dim Y As Integer
Dim X As Integer
Dim SaveAs As String
Dim SaveAs2 As String
Dim P1 As Boolean
Dim P2 As Boolean

P1 = False
P2 = False

              ' Reset Moficaciones
'              MapInfo(Map).Datos  - dat
'              MapInfo(Map).Bloqueos  - map
'              MapInfo(Map).Triggers - map
'              MapInfo(Map).NPCs - inf
'              MapInfo(Map).Objs - inf
'              MapInfo(Map).Telep - inf

Dim Por As String

If MapInfo(N).Bloqueos = True Or MapInfo(N).Triggers = True Then

    Por = IIf(MapInfo(N).Bloqueos = True, "Bloqueos,", "") & IIf(MapInfo(N).Triggers = True, "Triggers,", "")
    If Len(Por) > 3 Then
        Por = Left(Por, Len(Por) - 1)
        Por = Replace(Por, ",", "/")
    End If
    
    Call SendData(ToAdmins, 0, 0, "||Guardando Map" & N & ".map (Por cambios en " & Por & ")" & FONTTYPE_ADMIN)
    
    Dim Num1 As Long

    SaveAs = App.Path & "\WorldBackUP\Map" & N & ".map"
    If FileExist(SaveAs, vbNormal) Then
        Kill SaveAs
    End If
    
    'Open .map file
    Num1 = FreeFile
    Open SaveAs For Binary As Num1
    Seek Num1, 1
    
    'map Header
        
    Put Num1, , MapInfo(N).MapVersion
    Put Num1, , MiCabecera
    Put Num1, , TempINT
    Put Num1, , TempINT
    Put Num1, , TempINT
    Put Num1, , TempINT
    
    P1 = True
Else
    Call SendData(ToAdmins, 0, 0, "||No se detectaron cambios en Map" & N & ".map" & FONTTYPE_ADMIN)
End If

If MapInfo(N).NPCs = True Or MapInfo(N).Objs = True Or MapInfo(N).Telep = True Then

    
    Por = IIf(MapInfo(N).NPCs = True, "NPCs,", "") & IIf(MapInfo(N).Objs = True, "Objetos,", "") & IIf(MapInfo(N).Telep = True, "Teletransportes,", "")
    If Len(Por) > 3 Then
        Por = Left(Por, Len(Por) - 1)
        Por = Replace(Por, ",", "/")
    End If
    
    Call SendData(ToAdmins, 0, 0, "||Guardando Map" & N & ".inf (Por cambios en " & Por & ")" & FONTTYPE_ADMIN)
        
    Dim Num2 As Long

    SaveAs2 = App.Path & "\WorldBackUP\Map" & N & ".inf"
    If FileExist(SaveAs2, vbNormal) Then
        Kill SaveAs2
    End If
    
    'Open .inf file
    Num2 = FreeFile
    Open SaveAs2 For Binary As Num2
    Seek Num2, 1
    
    'inf Header
    Put Num2, , TempINT
    Put Num2, , TempINT
    Put Num2, , TempINT
    Put Num2, , TempINT
    Put Num2, , TempINT
    
    P2 = True
Else
    Call SendData(ToAdmins, 0, 0, "||No se detectaron cambios en Map" & N & ".inf" & FONTTYPE_ADMIN)
End If

If P1 = True Or P2 = True Then
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            If P1 = True Then
                '.map file
                Put Num1, , MapData(N, X, Y).Blocked
                
                For LoopC = 1 To 4
                    Put Num1, , MapData(N, X, Y).Graphic(LoopC)
                Next LoopC
                
                'Lugar vacio para futuras expansiones
                Put Num1, , MapData(N, X, Y).trigger
                
                Put Num1, , TempINT
            End If
            
            If P2 = True Then
                '.inf file
                'Tile exit
                Put Num2, , MapData(N, X, Y).TileExit.Map
                Put Num2, , MapData(N, X, Y).TileExit.X
                Put Num2, , MapData(N, X, Y).TileExit.Y
                
                'NPC
                If MapData(N, X, Y).NpcIndex > 0 Then
                    Put Num2, , Npclist(MapData(N, X, Y).NpcIndex).Numero
                Else
                    Put Num2, , 0
                End If
                'Object
                
                If MapData(N, X, Y).OBJInfo.ObjIndex > 0 Then
                    If ObjData(MapData(N, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_FOGATA Then
                        MapData(N, X, Y).OBJInfo.ObjIndex = 0
                        MapData(N, X, Y).OBJInfo.Amount = 0
                    End If
        '            If ObjData(MapData(n, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_MANCHAS Then
        '                MapData(n, X, Y).OBJInfo.ObjIndex = 0
        '                MapData(n, X, Y).OBJInfo.Amount = 0
        '            End If
                End If
                
                Put Num2, , MapData(N, X, Y).OBJInfo.ObjIndex
                Put Num2, , MapData(N, X, Y).OBJInfo.Amount
                
                'Empty place holders for future expansion
                Put Num2, , TempINT
                Put Num2, , TempINT
            End If
        Next X
    Next Y
End If


If P1 = True Then
    'Close .map file
    Close Num1
    P1 = False
    MapInfo(N).Triggers = False
    MapInfo(N).Bloqueos = False
End If

If P2 = True Then
    'Close .inf file
    Close Num2
    P2 = False
    MapInfo(N).Telep = False
    MapInfo(N).NPCs = False
    MapInfo(N).Objs = False
End If

If MapInfo(N).Datos = True Then

    Call SendData(ToAdmins, 0, 0, "||Guardando Map" & N & ".dat (Por cambios de configuración)" & FONTTYPE_ADMIN)

    'write .dat file
    SaveAs = App.Path & "\WorldBackUP\Map" & N & ".dat"
    Call WriteVar(SaveAs, "Mapa" & N, "Name", MapInfo(N).Name)
    Call WriteVar(SaveAs, "Mapa" & N, "MusicNum", MapInfo(N).Music)
    Call WriteVar(SaveAs, "Mapa" & N, "StartPos", MapInfo(N).StartPos.Map & "-" & MapInfo(N).StartPos.X & "-" & MapInfo(N).StartPos.Y)
    
    Call WriteVar(SaveAs, "Mapa" & N, "Terreno", MapInfo(N).Terreno)
    Call WriteVar(SaveAs, "Mapa" & N, "Zona", MapInfo(N).Zona)
    Call WriteVar(SaveAs, "Mapa" & N, "Restringir", MapInfo(N).Restringir)
    Call WriteVar(SaveAs, "Mapa" & N, "BackUp", str(MapInfo(N).BackUp))
    
    Call WriteVar(SaveAs, "Mapa" & N, "MagiaSinEfecto", str(MapInfo(N).MagiaSinEfecto))
    
    If MapInfo(N).Pk Then
        Call WriteVar(SaveAs, "Mapa" & N, "pk", "0")
    Else
        Call WriteVar(SaveAs, "Mapa" & N, "pk", "1")
    End If
    
    MapInfo(N).Datos = False
Else
    Call SendData(ToAdmins, 0, 0, "||No se detectaron cambios en Map" & N & ".dat" & FONTTYPE_ADMIN)
End If

Exit Sub
FalloMapa:
    MsgBox Err.Number & " - " & Err.Description

End Sub

Sub LoadArmasHerreria()

Dim N As Integer, lC As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lC = 1 To N
    ArmasHerrero(lC) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lC, "Index"))
Next lC


End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lC As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lC = 1 To N
    ArmadurasHerrero(lC) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lC, "Index"))
Next lC

End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lC As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lC = 1 To N
    ObjCarpintero(lC) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lC, "Index"))
Next lC

End Sub




Sub LoadOBJData_Nuevo()
Dim ErrorXF As String
'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler
'On Error GoTo 0
ErrorXF = "Iniciando"
If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsLeerInis
Dim Tiempo As Long

Tiempo = GetTickCount

' 577 = 100
' 570 =
frmCargando.Label1(0).Caption = "Preparando..."
ErrorXF = "Leyendo obj.dat"
Leer.Abrir DatPath & "Obj.dat"
'j = val(Leer.DarValor("INIT", "NumObjs"))  '

'obtiene el numero de obj
NumObjDatas = val(Leer.DarValor("INIT", "NumObjs"))

frmCargando.Cargar.Value = 0
frmCargando.Cargar.Min = 0
frmCargando.Cargar.max = NumObjDatas
frmGeneral.ProG1.Min = 0
frmGeneral.ProG1.max = NumObjDatas
frmGeneral.ProG1.Value = 0
frmGeneral.ProG1.Visible = True

ErrorXF = "Inicializando variables"

ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo nombre"
    ObjData(Object).Name = Leer.DarValor("OBJ" & Object, "Name")
    
'    ' [GS] Intercambio
'    ObjData(Object).NumIntercambio = INIDarClaveInt(A, S, "NumIntercambio")
'    If ObjData(Object).NumIntercambio > 0 And ObjData(Object).NumIntercambio < 11 Then
'        For TT = 1 To ObjData(Object).NumIntercambio
'            ObjData(Object).Intercambio(TT) = INIDarClaveStr(A, S, "Intercambio" & TT)
'        Next
'        ' Anti Errores
'        For TT = 1 To ObjData(Object).NumIntercambio
'            If IsNumeric(ReadField(1, ObjData(Object).Intercambio(TT), Asc("-"))) And IsNumeric(ReadField(2, ObjData(Object).Intercambio(TT), Asc("-"))) Then
'                ObjData(Object).NumIntercambio = 0
'                ObjData(Object).CantIntercambia = 0
'                Exit For
'            End If
'        Next
'        ObjData(Object).CantIntercambia = INIDarClaveInt(A, S, "CantInt")
'        If ObjData(Object).CantIntercambia <= 0 Or ObjData(Object).CantIntercambia > 10000 Then
'            ObjData(Object).NumIntercambio = 0
'            ObjData(Object).CantIntercambia = 0
'            Exit Sub
'        End If
'    Else
'        ObjData(Object).NumIntercambio = 0
'        ObjData(Object).CantIntercambia = 0
'    End If
'    ' [/GS]
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo GrhIndex"
    
    ObjData(Object).GrhIndex = val(Leer.DarValor("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - ObjType y SubTipo"
    
    ObjData(Object).ObjType = val(Leer.DarValor("OBJ" & Object, "ObjType"))
    ObjData(Object).SubTipo = val(Leer.DarValor("OBJ" & Object, "Subtipo"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Newbie"
    
    ObjData(Object).Newbie = val(Leer.DarValor("OBJ" & Object, "Newbie"))

    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo DefensaMagicaMin/Max"

    ObjData(Object).DefensaMagicaMax = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMin"))
    If ObjData(Object).DefensaMagicaMin > ObjData(Object).DefensaMagicaMax Then ObjData(Object).DefensaMagicaMin = ObjData(Object).DefensaMagicaMax - 1

    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Escudo"
    
    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
        ObjData(Object).ShieldAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
        If val(Leer.DarValor("OBJ" & Object, "Devuelve")) <= 100 Then
            ObjData(Object).Devuelve = val(Leer.DarValor("OBJ" & Object, "Devuelve"))
            ' [GS] Devuelve
            If ObjData(Object).Devuelve < 0 Then ObjData(Object).Devuelve = 0
            If ObjData(Object).Devuelve > 100 Then ObjData(Object).Devuelve = 100
            ' [/GS]
        End If
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Casco"
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
        ObjData(Object).CascoAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Numero de Ropaje/HechizoIndex"
    
    ObjData(Object).Ropaje = val(Leer.DarValor("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.DarValor("OBJ" & Object, "HechizoIndex"))
      
      
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Numero de Cabeza de Ropaje"
      ' [GS] Cabeza
    ObjData(Object).Cabeza = val(Leer.DarValor("OBJ" & Object, "NumCabeza"))
    ' [/GS]
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo MinNivel"
    ' [GS]
    ObjData(Object).MinNivel = val(Leer.DarValor("OBJ" & Object, "MinNivel"))
    ' [/GS]
    If ObjData(Object).MinNivel > STAT_MAXELV Then
        Call Alerta("El Objeto " & Object & " tiene un MinNivel superior al Nivel Maximo.")
    ElseIf ObjData(Object).MinNivel < 1 Then
        ObjData(Object).MinNivel = 0
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Consumo de Mana"
    ' [GS] Objeto consume mana?
    ObjData(Object).mana = val(Leer.DarValor("OBJ" & Object, "Mana"))
    ' [/GS]
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Arma"
    
    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
            ObjData(Object).WeaponAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.DarValor("OBJ" & Object, "Apuñala"))
            ObjData(Object).Paraliza = val(Leer.DarValor("OBJ" & Object, "Paraliza"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = val(Leer.DarValor("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.DarValor("OBJ" & Object, "Municiones"))
            ' [GS]
            ObjData(Object).SoloNPC = val(Leer.DarValor("OBJ" & Object, "SoloNPC"))
            ObjData(Object).Paraliza = val(Leer.DarValor("OBJ" & Object, "Paraliza"))
            If ObjData(Object).Paraliza > 100 Then ObjData(Object).Paraliza = 100
            If ObjData(Object).Paraliza < 1 Then ObjData(Object).Paraliza = 1
            ' [/GS]
            ' [GS] Dos manos?
            ObjData(Object).DosManos = val(Leer.DarValor("OBJ" & Object, "DosManos"))
            ' [/GS]
            ObjData(Object).StaffPower = val(Leer.DarValor("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.DarValor("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.DarValor("OBJ" & Object, "Refuerzo"))
    End If

    ' [GS] Magic
    If MatematicasConComa = True Then
        ObjData(Object).Magic = Replace(Leer.DarValor("OBJ" & Object, "Magic"), ".", ",")
        ObjData(Object).Poder = Replace(Leer.DarValor("OBJ" & Object, "Poder"), ".", ",")
        ObjData(Object).Agilidad = Replace(Leer.DarValor("OBJ" & Object, "Agilidad"), ".", ",")
    Else
        ObjData(Object).Magic = Replace(Leer.DarValor("OBJ" & Object, "Magic"), ",", ".")
        ObjData(Object).Poder = Replace(Leer.DarValor("OBJ" & Object, "Poder"), ",", ".")
        ObjData(Object).Agilidad = Replace(Leer.DarValor("OBJ" & Object, "Agilidad"), ",", ".")
    End If
    ' [/GS]
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Armadura"
    
    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            ' [GS] Devuelve
            If val(Leer.DarValor("OBJ" & Object, "Devuelve")) <= 100 Then
                ObjData(Object).Devuelve = val(Leer.DarValor("OBJ" & Object, "Devuelve"))
                If ObjData(Object).Devuelve < 0 Then ObjData(Object).Devuelve = 0
                If ObjData(Object).Devuelve > 100 Then ObjData(Object).Devuelve = 100
                ' [/GS]
            End If
            ' [GS] No paralisis
            ObjData(Object).NoParalisis = val(Leer.DarValor("OBJ" & Object, "NoParalisis"))
            If ObjData(Object).NoParalisis < 0 Then ObjData(Object).NoParalisis = 0
            If ObjData(Object).NoParalisis > 100 Then ObjData(Object).NoParalisis = 100
            ' [/GS]
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Herramienta"
    
    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedades de Instrumento"
    
    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
        ObjData(Object).Snd1 = val(Leer.DarValor("OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = val(Leer.DarValor("OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = val(Leer.DarValor("OBJ" & Object, "SND3"))
        ObjData(Object).MinInt = val(Leer.DarValor("OBJ" & Object, "MinInt"))
    End If
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Informacion de Herreria"
    
    ObjData(Object).LingoteIndex = val(Leer.DarValor("OBJ" & Object, "LingoteIndex"))
    
    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
        ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
    End If
    
    ObjData(Object).MineralIndex = val(Leer.DarValor("OBJ" & Object, "MineralIndex"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo HP"
    
    ObjData(Object).MaxHP = val(Leer.DarValor("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.DarValor("OBJ" & Object, "MinHP"))
  
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Genero Requerido"
    
    ObjData(Object).MUJER = val(Leer.DarValor("OBJ" & Object, "Mujer"))
    ObjData(Object).HOMBRE = val(Leer.DarValor("OBJ" & Object, "Hombre"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Hambre y Agua"
    
    ObjData(Object).MinHam = val(Leer.DarValor("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.DarValor("OBJ" & Object, "MinAgu"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Defensa"
    
    ObjData(Object).MinDef = val(Leer.DarValor("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.DarValor("OBJ" & Object, "MAXDEF"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Si tiene Respawn"
    
    ObjData(Object).Respawn = val(Leer.DarValor("OBJ" & Object, "ReSpawn"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Si es de Raza Enana"
    
    ObjData(Object).RazaEnana = val(Leer.DarValor("OBJ" & Object, "RazaEnana"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Valor"
    
    ObjData(Object).Valor = val(Leer.DarValor("OBJ" & Object, "Valor"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo si es Crucial"
    
    ObjData(Object).Crucial = val(Leer.DarValor("OBJ" & Object, "Crucial"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo informacion de Puerta"
    
    ObjData(Object).Cerrada = val(Leer.DarValor("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.DarValor("OBJ" & Object, "Llave"))
            ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    End If


    ErrorXF = "Objeto" & NumObjDatas & " - Analisis anti lageadores"
    
    ' [GS] No chiteros lageadores
    If ObjData(Object).Valor <= 0 Then
        ObjData(Object).Valor = 0
        ObjData(Object).NoSeVende = 1
    End If
    ' [/GS]

    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
        ObjData(Object).IndexAbierta = val(Leer.DarValor("OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerrada = val(Leer.DarValor("OBJ" & Object, "IndexCerrada"))
        ObjData(Object).IndexCerradaLlave = val(Leer.DarValor("OBJ" & Object, "IndexCerradaLlave"))
    End If
    
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo puertas y llaves"
    
    'Puertas y llaves
    ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    
    ObjData(Object).Texto = Leer.DarValor("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.DarValor("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.DarValor("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.DarValor("OBJ" & Object, "ID")
    
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Clases Prohibidas"
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Clase2Num(Leer.DarValor("OBJ" & Object, "CP" & i))
    Next
            
    ObjData(Object).Resistencia = val(Leer.DarValor("OBJ" & Object, "Resistencia"))
    
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Pociones"

    'Pociones
    If ObjData(Object).ObjType = 11 Then
        ObjData(Object).TipoPocion = val(Leer.DarValor("OBJ" & Object, "TipoPocion"))
        ObjData(Object).MaxModificador = val(Leer.DarValor("OBJ" & Object, "MaxModificador"))
        ObjData(Object).MinModificador = val(Leer.DarValor("OBJ" & Object, "MinModificador"))
        ObjData(Object).DuracionEfecto = val(Leer.DarValor("OBJ" & Object, "DuracionEfecto"))
    End If

    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo requisitos de Carpinteria"
    
    ObjData(Object).SkCarpinteria = val(Leer.DarValor("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.DarValor("OBJ" & Object, "Madera"))
    
    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.DarValor("OBJ" & Object, "Paraliza"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.DarValor("OBJ" & Object, "MinST"))

    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo si es Vendible/SeCae o no"
    
    ' [GS]
    ObjData(Object).NoSeCae = val(Leer.DarValor("OBJ" & Object, "NoSeCae"))
    ObjData(Object).NoSeVende = val(Leer.DarValor("OBJ" & Object, "NoSeVende"))
    ' [/GS]
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Propiedad de NO SE PASA"
    ObjData(Object).NoSePasa = val(Leer.DarValor("OBJ" & Object, "NoSePasa"))
    
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo Clase Exclusiva"
    
    ' [GS]
    ObjData(Object).ExclusivoClase = Clase2Num(Leer.DarValor("OBJ" & Object, "ExclusivoClase"))
    ' [/GS]
        
    If frmCargando.Visible Then
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
        frmCargando.Label1(0).Caption = ObjData(Object).Name
    End If
    If frmGeneral.Visible Then frmGeneral.ProG1.Value = frmGeneral.ProG1.Value + 1
    ErrorXF = "Objeto" & NumObjDatas & " - Leyendo el siguiente objeto :S"
Next Object
frmGeneral.ProG1.Visible = False
ErrorXF = "Fin de lectura"
Tiempo = GetTickCount - Tiempo
Call LogCOSAS("Tiempos", "Cargado de Objetos " & str(Tiempo / 1000) & " segundos.", False)
frmCargando.Label1(0).Caption = "Haciendo copia de seguridad"
ErrorXF = "Haciendo copia de seguridad"
Call HacerDAT("Obj.dat")
frmCargando.Label1(0).Caption = ""
ErrorXF = "OK"
Exit Sub

errhandler:

Call LogError("Error en LoadOBJData_Nuevo: " & Err.Number & " " & Err.Description & " Error en: " & ErrorXF)
MsgBox "Error en el Cargado de Objetos" & vbCrLf & "Error: " & Err.Number & " - " & Err.Description & vbCrLf & "ERROR EN " & ErrorXF, vbCritical
Call RepararDAT("Obj.dat", "Obj_2.dat")




End Sub


Sub LoadUserStats(UserIndex As Integer, UserFile As String)



Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = GetVar(UserFile, "ATRIBUTOS", "AT" & LoopC)
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = val(GetVar(UserFile, "SKILLS", "SK" & LoopC))
Next

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = val(GetVar(UserFile, "Hechizos", "H" & LoopC))
Next

UserList(UserIndex).Stats.GLD = val(GetVar(UserFile, "STATS", "GLD"))
UserList(UserIndex).Stats.banco = val(GetVar(UserFile, "STATS", "BANCO"))

UserList(UserIndex).Stats.MET = val(GetVar(UserFile, "STATS", "MET"))
UserList(UserIndex).Stats.MaxHP = val(GetVar(UserFile, "STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = val(GetVar(UserFile, "STATS", "MinHP"))

UserList(UserIndex).Stats.FIT = val(GetVar(UserFile, "STATS", "FIT"))
UserList(UserIndex).Stats.MinSta = val(GetVar(UserFile, "STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = val(GetVar(UserFile, "STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = val(GetVar(UserFile, "STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = val(GetVar(UserFile, "STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = val(GetVar(UserFile, "STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = val(GetVar(UserFile, "STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = val(GetVar(UserFile, "STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = val(GetVar(UserFile, "STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = val(GetVar(UserFile, "STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = val(GetVar(UserFile, "STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = val(GetVar(UserFile, "STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.exp = val(GetVar(UserFile, "STATS", "EXP"))
UserList(UserIndex).Stats.ELU = val(GetVar(UserFile, "STATS", "ELU"))
UserList(UserIndex).Stats.ELV = val(GetVar(UserFile, "STATS", "ELV"))

' [GS] Party
UserList(UserIndex).flags.LiderParty = 0
' [/GS]

UserList(UserIndex).Stats.UsuariosMatados = val(GetVar(UserFile, "MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.CriminalesMatados = val(GetVar(UserFile, "MUERTES", "CrimMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = val(GetVar(UserFile, "MUERTES", "NpcsMuertes"))

' 0.12b1
UserList(UserIndex).flags.PertAlCons = IIf(val(GetVar(UserFile, "CONSEJO", "PERTENECE")) = 1, True, False)
UserList(UserIndex).flags.PertAlConsCaos = IIf(val(GetVar(UserFile, "CONSEJO", "PERTENECECAOS")) = 1, True, False)




End Sub

Sub LoadUserReputacion(UserIndex As Integer, UserFile As String)

UserList(UserIndex).Reputacion.AsesinoRep = val(GetVar(UserFile, "REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(GetVar(UserFile, "REP", "Dandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(GetVar(UserFile, "REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(GetVar(UserFile, "REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(GetVar(UserFile, "REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(GetVar(UserFile, "REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(GetVar(UserFile, "REP", "Promedio"))

End Sub


Sub LoadUserInit(UserIndex As Integer, UserFile As String)


Dim LoopC As Integer
Dim ln As String
Dim ln2 As String

' [GS]
UserList(UserIndex).flags.BorrarAlSalir = False
UserList(UserIndex).flags.PocionRepelente = False
UserList(UserIndex).flags.TieneMensaje = False
UserList(UserIndex).Administracion.Activado = False
UserList(UserIndex).flags.UltimoNickColor = ""
UserList(UserIndex).Silenciado = False
UserList(UserIndex).NoExiste = False
' [/GS]
'[Sicarul] Hiper-AO
UserList(UserIndex).flags.Casado = GetVar(UserFile, "FLAGS", "Casado")
If UserList(UserIndex).flags.Casado = "0" Then UserList(UserIndex).flags.Casado = ""
'[/Sicarul]
UserList(UserIndex).Faccion.ArmadaReal = val(GetVar(UserFile, "FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = val(GetVar(UserFile, "FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = val(GetVar(UserFile, "FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = val(GetVar(UserFile, "FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(GetVar(UserFile, "FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = val(GetVar(UserFile, "FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(GetVar(UserFile, "FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = val(GetVar(UserFile, "FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = val(GetVar(UserFile, "FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = val(GetVar(UserFile, "FACCIONES", "recReal"))

' 0.12b1
UserList(UserIndex).Faccion.Reenlistadas = val(GetVar(UserFile, "FACCIONES", "Reenlistadas"))

UserList(UserIndex).flags.YaVoto = False

UserList(UserIndex).flags.Muerto = val(GetVar(UserFile, "FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = val(GetVar(UserFile, "FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = val(GetVar(UserFile, "FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = val(GetVar(UserFile, "FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = val(GetVar(UserFile, "FLAGS", "Desnudo"))

UserList(UserIndex).flags.TiempoOnline = val(GetVar(UserFile, "FLAGS", "TiempoOnline"))
If UserList(UserIndex).flags.TiempoOnline < 0 Then UserList(UserIndex).flags.TiempoOnline = 0
If UserList(UserIndex).flags.TiempoOnline > tLong Then UserList(UserIndex).flags.TiempoOnline = 0

UserList(UserIndex).flags.Envenenado = val(GetVar(UserFile, "FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = val(GetVar(UserFile, "FLAGS", "Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If
UserList(UserIndex).flags.Navegando = val(GetVar(UserFile, "FLAGS", "Navegando"))


UserList(UserIndex).Counters.Pena = val(GetVar(UserFile, "COUNTERS", "Pena"))

UserList(UserIndex).Email = GetVar(UserFile, "CONTACTO", "Email")

UserList(UserIndex).genero = Gen2Num(GetVar(UserFile, "INIT", "Genero"))
UserList(UserIndex).clase = Clase2Num(GetVar(UserFile, "INIT", "Clase"))
UserList(UserIndex).raza = Raza2Num(GetVar(UserFile, "INIT", "Raza"))
UserList(UserIndex).Hogar = GetVar(UserFile, "INIT", "Hogar")
UserList(UserIndex).Char.Heading = val(GetVar(UserFile, "INIT", "Heading"))

' [GS] Counter mode
UserList(UserIndex).flags.CS_Esta = False
' [/GS]

' [GS] Aventura
UserList(UserIndex).flags.AV_Lugar = GetVar(UserFile, "AVENTURA", "IniPos")
UserList(UserIndex).flags.AV_Tiempo = val(GetVar(UserFile, "AVENTURA", "Tiempo"))
' [/GS]

UserList(UserIndex).flags.UltimoMensaje = 0
UserList(UserIndex).flags.UltimoEST = ""

UserList(UserIndex).OrigChar.Head = val(GetVar(UserFile, "INIT", "Head"))
UserList(UserIndex).OrigChar.Body = val(GetVar(UserFile, "INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = val(GetVar(UserFile, "INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = val(GetVar(UserFile, "INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = val(GetVar(UserFile, "INIT", "Casco"))
UserList(UserIndex).OrigChar.Heading = SOUTH

If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


UserList(UserIndex).desc = GetVar(UserFile, "INIT", "Desc")


UserList(UserIndex).Pos.Map = val(ReadField(1, GetVar(UserFile, "INIT", "Position"), 45))
' [GS] es aventura
If UserList(UserIndex).Pos.Map = MapaAventura And UserList(UserIndex).flags.AV_Tiempo < 0 Then
    UserList(UserIndex).flags.AV_Esta = True
Else
    UserList(UserIndex).flags.AV_Esta = False
End If
' [/GS]
UserList(UserIndex).Pos.X = val(ReadField(2, GetVar(UserFile, "INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = val(ReadField(3, GetVar(UserFile, "INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = GetVar(UserFile, "Inventory", "CantidadItems")

Dim loopd As Integer

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = val(GetVar(UserFile, "BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    ln2 = GetVar(UserFile, "BancoInventory", "Obj" & loopd)
    UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex = val(ReadField(1, ln2, 45))
    UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
Next loopd
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(UserFile, "Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = val(GetVar(UserFile, "Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = val(GetVar(UserFile, "Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Cabeza <> 0 Then
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Cabeza > 0 Then
            UserList(UserIndex).Char.Head = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Cabeza
        ElseIf ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Cabeza = -1 Then
            UserList(UserIndex).Char.Head = 0
        End If
    End If
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = val(GetVar(UserFile, "Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = val(GetVar(UserFile, "Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = val(GetVar(UserFile, "Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = val(GetVar(UserFile, "Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(UserIndex).Invent.HerramientaEqpSlot = val(GetVar(UserFile, "Inventory", "HerramientaSlot"))
If UserList(UserIndex).Invent.HerramientaEqpSlot > 0 Then
    UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex
End If

'0.12b1
UserList(UserIndex).Invent.Accesorio1EqpSlot = 0
UserList(UserIndex).Invent.Accesorio2EqpSlot = 0

' [GS] Accesorios
UserList(UserIndex).Invent.Accesorio1EqpSlot = val(GetVar(UserFile, "Inventory", "Accesorio1Slot"))
If UserList(UserIndex).Invent.Accesorio1EqpSlot > 0 Then
    UserList(UserIndex).Invent.Accesorio1EqpObjIndex = val(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.Accesorio1EqpSlot).ObjIndex)
    If UserList(UserIndex).Invent.Accesorio1EqpObjIndex <= 0 Then
        UserList(UserIndex).Invent.Accesorio1EqpObjIndex = 0
        UserList(UserIndex).Invent.Accesorio1EqpSlot = 0
    End If
End If

UserList(UserIndex).Invent.Accesorio2EqpSlot = val(GetVar(UserFile, "Inventory", "Accesorio2Slot"))
If UserList(UserIndex).Invent.Accesorio2EqpSlot > 0 Then
    UserList(UserIndex).Invent.Accesorio2EqpObjIndex = val(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.Accesorio2EqpSlot).ObjIndex)
    If UserList(UserIndex).Invent.Accesorio2EqpObjIndex <= 0 Then
        UserList(UserIndex).Invent.Accesorio2EqpObjIndex = 0
        UserList(UserIndex).Invent.Accesorio2EqpSlot = 0
    End If
End If

' [/GS]

' [GS] Administracion ??

Call CargarADMIN(UserFile, UserIndex)

' [/GS]

UserList(UserIndex).NroMacotas = val(GetVar(UserFile, "Mascotas", "NroMascotas"))

'Lista de objetos
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasType(LoopC) = val(GetVar(UserFile, "Mascotas", "Mas" & LoopC))
Next LoopC

UserList(UserIndex).GuildInfo.FundoClan = val(GetVar(UserFile, "Guild", "FundoClan"))
UserList(UserIndex).GuildInfo.EsGuildLeader = val(GetVar(UserFile, "Guild", "EsGuildLeader"))
UserList(UserIndex).GuildInfo.Echadas = val(GetVar(UserFile, "Guild", "Echadas"))
UserList(UserIndex).GuildInfo.Solicitudes = val(GetVar(UserFile, "Guild", "Solicitudes"))
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(GetVar(UserFile, "Guild", "SolicitudesRechazadas"))
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(GetVar(UserFile, "Guild", "VecesFueGuildLeader"))
UserList(UserIndex).GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
UserList(UserIndex).GuildInfo.ClanesParticipo = val(GetVar(UserFile, "Guild", "ClanesParticipo"))
UserList(UserIndex).GuildInfo.GuildPoints = val(GetVar(UserFile, "Guild", "GuildPts"))

UserList(UserIndex).GuildInfo.ClanFundado = GetVar(UserFile, "Guild", "ClanFundado")
UserList(UserIndex).GuildInfo.GuildName = GetVar(UserFile, "Guild", "GuildName")

End Sub

Sub CargarADMIN(ByVal UserFile As String, ByVal UserIndex As Integer)
On Error Resume Next
Dim LoopC As Integer
' [GS] Administracion ??

UserList(UserIndex).Administracion.EnPrueba = True
UserList(UserIndex).Administracion.Activado = False

LoopC = val(GetVar(UserFile, "ADMINISTRACION", "Activado"))
If LoopC = 1 Then
    LoopC = val(GetVar(UserFile, "ADMINISTRACION", "EnPrueba"))
    If LoopC = 1 Then
        UserList(UserIndex).Administracion.EnPrueba = True
    Else
        UserList(UserIndex).Administracion.EnPrueba = False
    End If
    UserList(UserIndex).Administracion.Config = GetVar(UserFile, "ADMINISTRACION", "Config")
    UserList(UserIndex).Administracion.Activado = True
    UserList(UserIndex).Administracion.MaxCP = val(GetVar(UserFile, "ADMINISTRACION", "CP"))
    For LoopC = 1 To UserList(UserIndex).Administracion.MaxCP
        UserList(UserIndex).Administracion.CP(LoopC) = GetVar(UserFile, "ADMINISTRACION", "CP" & LoopC)
    Next
Else
    UserList(UserIndex).Administracion.Activado = False
End If

' [/GS]
End Sub




Function GetVar(file As String, Main As String, Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function


Sub CargarBackUp()

'Call LogTarea("Sub CargarBackUp")
On Error GoTo MAN
Dim ParteError As Integer
ParteError = 0
If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando backup."
Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempINT As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim c$
Dim Tiempo As Long
Dim Leer As New clsLeerInis
Dim NoSirve As Boolean

Tiempo = GetTickCount

Leer.Abrir DatPath & "Map.dat"
NumMaps = val(Leer.DarValor("INIT", "NumMaps"))
MapPath = Leer.DarValor("INIT", "MapPath")

frmCargando.Cargar.Min = 0
frmCargando.Cargar.max = NumMaps
frmCargando.Cargar.Value = 0
frmGeneral.ProG1.Min = 0
frmGeneral.ProG1.max = NumMaps
frmGeneral.ProG1.Value = 0
frmGeneral.ProG1.Visible = True
frmCargando.Label1(0).Caption = "Preparando..."

' [GS] Correccion para mapa pretoriano
If MAPA_PRETORIANO <> 0 Then
    'MsgBox "Activado Sistema Pretoriano, se cargaran todos los mapas hasta " & IIf(NumMaps < MAPA_PRETORIANO, MAPA_PRETORIANO, NumMaps) & "."
    If NumMaps < MAPA_PRETORIANO Then
        NumMaps = MAPA_PRETORIANO
        frmCargando.Cargar.max = NumMaps
        frmGeneral.ProG1.max = NumMaps
    End If
End If
' [/GS]


ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    MapInfo(Map).Cargado = False
    NoSirve = False
    NoSirve = val(Leer.DarValor("MAP" & Map, "NoCargar"))
    If NoSirve = False Then
        FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
        
        If FileExist(FileNamE, vbNormal) Then
            Open App.Path & "\WorldBackUp\Map" & Map & ".map" For Binary As #1
        Else
            Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        End If
        
        If FileExist(App.Path & "\WorldBackUp\Map" & Map & ".inf", vbNormal) Then
            Open App.Path & "\WorldBackUp\Map" & Map & ".inf" For Binary As #2
        Else
            Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        End If
        
        If FileExist(App.Path & "\WorldBackUp\Map" & Map & ".dat", vbNormal) Then
            c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
        Else
            c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
        End If
        
            Seek #1, 1
            Seek #2, 1
            'map Header
            Get #1, , MapInfo(Map).MapVersion
            Get #1, , MiCabecera
            Get #1, , TempINT
            Get #1, , TempINT
            Get #1, , TempINT
            Get #1, , TempINT
            'inf Header
            Get #2, , TempINT
            Get #2, , TempINT
            Get #2, , TempINT
            Get #2, , TempINT
            Get #2, , TempINT
            'Load arrays
                        DoEvents
            For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
                        '.dat file
                        Get #1, , MapData(Map, X, Y).Blocked
                        
                        'Get GRH number
                        For LoopC = 1 To 4
                            Get #1, , MapData(Map, X, Y).Graphic(LoopC)
                        Next LoopC
                        
                        'Space holder for future expansion
                        Get #1, , MapData(Map, X, Y).trigger
                        Get #1, , TempINT
                        
                                            
                        '.inf file
                        Get #2, , MapData(Map, X, Y).TileExit.Map
                        Get #2, , MapData(Map, X, Y).TileExit.X
                        Get #2, , MapData(Map, X, Y).TileExit.Y
                        
                        'Get and make NPC
                        Get #2, , MapData(Map, X, Y).NpcIndex
                        If MapData(Map, X, Y).NpcIndex > 0 Then
                        
                            ParteError = 1
                            MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                            'Si el npc debe hacer respawn en la pos
                            'original la guardamos
                            ParteError = 0
                            
                            If Npclist(MapData(Map, X, Y).NpcIndex).Numero > 499 Then
                                npcfile = DatPath & "NPCs-HOSTILES.dat"
                            Else
                                npcfile = DatPath & "NPCs.dat"
                            End If
                            
                            Dim fl As Byte
                            fl = val(GetVar(npcfile, "NPC" & Npclist(MapData(Map, X, Y).NpcIndex).Numero, "PosOrig"))
                            If fl = 1 Then
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                            Else
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = 0
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = 0
                                Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = 0
                            End If
            
                            Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                            Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                            Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                            
    
                            
                            'Si existe el backup lo cargamos
                            If Npclist(MapData(Map, X, Y).NpcIndex).flags.BackUp = 1 Then
                                    'cargamos el nuevo del backup
                                    Call CargarNpcBackUp(MapData(Map, X, Y).NpcIndex, Npclist(MapData(Map, X, Y).NpcIndex).Numero)
                                    
                            End If
                            
                            Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                        End If
    
                        'Get and make Object
                        Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
                        Get #2, , MapData(Map, X, Y).OBJInfo.Amount
            
                        'Space holder for future expansion (Objects, ect.
                        Get #2, , DummyInt
                        Get #2, , DummyInt
                Next X
            Next Y
            Close #1
            Close #2
              MapInfo(Map).Name = GetVar(c$, "Mapa" & Map, "Name")
              MapInfo(Map).Music = GetVar(c$, "Mapa" & Map, "MusicNum")
              MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
              MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
              MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
              If val(GetVar(c$, "Mapa" & Map, "Pk")) = 0 Then
                    MapInfo(Map).Pk = True
              Else
                    MapInfo(Map).Pk = False
              End If
              MapInfo(Map).Restringir = GetVar(c$, "Mapa" & Map, "Restringir")
              MapInfo(Map).BackUp = val(GetVar(c$, "Mapa" & Map, "BackUp"))
              MapInfo(Map).Terreno = GetVar(c$, "Mapa" & Map, "Terreno")
              MapInfo(Map).Zona = GetVar(c$, "Mapa" & Map, "Zona")
              MapInfo(Map).MagiaSinEfecto = val(GetVar(c$, "Mapa" & Map, "MagiaSinEfecto"))
              
              ' Reset Moficaciones
              MapInfo(Map).Datos = False
              MapInfo(Map).Bloqueos = False
              MapInfo(Map).NPCs = False
              MapInfo(Map).Objs = False
              MapInfo(Map).Triggers = False
              MapInfo(Map).Telep = False
          
          MapInfo(Map).Cargado = True
    End If
          
        If frmCargando.Visible Then
            frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
            frmCargando.Label1(0).Caption = "Mapa " & str(Map) & " de " & str(NumMaps)
            frmCargando.SetFocus
        End If
        If frmGeneral.Visible Then
            frmGeneral.ProG1.Value = frmGeneral.ProG1.Value + 1
            frmGeneral.Estado.SimpleText = "Cargando backup. " & str(Map) & " de " & str(NumMaps)
        End If
        DoEvents

Next Map



Tiempo = GetTickCount - Tiempo
Call LogCOSAS("Tiempos", "Cargado de Mapas de BackUP (" & NumMaps & ") " & str(Tiempo / 1000) & " segundos.", False)


frmGeneral.ProG1.Visible = False
FrmStat.Visible = False

Exit Sub

MAN:
If ParteError = 0 Then ' Error en el mapa
    Call MsgBox("Error durante la carga de mapas (de backup). El mapa " & Map & " posiblemente contiene errores.")
    Call LogError(Date & " ERROR EN MAPA BACKUP - Nro. " & Err.Number & " - " & Err.Description)
    If FileExist(App.Path & "\WorldBackUp\Seguridad\", vbDirectory) = True Then
        If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".dat", vbArchive) = True Then
            If FileExist(App.Path & "\WorldBackUp\map" & Map & ".dat", vbArchive) = False Then
                Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".dat", App.Path & "\WorldBackUp\map" & Map & ".dat")
            End If
        End If
        If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".map", vbArchive) = True And FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".inf", vbArchive) = True And FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".dat", vbArchive) = True Then
            Call BorrarArchivo(App.Path & "\WorldBackUp\map" & Map & ".map")
            Call BorrarArchivo(App.Path & "\WorldBackUp\map" & Map & ".inf")
            Call BorrarArchivo(App.Path & "\WorldBackUp\map" & Map & ".dat")
            Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".map", App.Path & "\WorldBackUp\map" & Map & ".map")
            Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".inf", App.Path & "\WorldBackUp\map" & Map & ".inf")
            Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".dat", App.Path & "\WorldBackUp\map" & Map & ".dat")
            If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".map", vbArchive) = True And FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".inf", vbArchive) = True And FileExist(App.Path & "\WorldBackUp\Seguridad\map" & Map & ".dat", vbArchive) = True Then
                MsgBox "Reparado con exito. Vuelva a ejecutar el servidor.", vbInformation
                End
            Else
                MsgBox "Ocurrio un error durante la reparacion, un archivo esta dañado. Pruebe repararlo manualmente.", vbCritical
                End
            End If
            DoEvents
        Else
            MsgBox "Imposible reparar el mapa " & Map & ", no hay copia de seguridad. Intentelo manualmente o eliminelo.", vbCritical
            End
        End If
    End If
Else
    Call MsgBox("Error durante la carga de mapas (de backup). El mapa " & Map & " contiene errores en referencia a los NPCs.")
    Call LogError(Date & " ERROR EN NPC BACKUP - Nro. " & Err.Number & " - " & Err.Description)
End If
' [GS]
End ' Minimo error y no abrimos el Host
' [/GS]
End Sub

Sub LoadMapData()


'Call LogTarea("Sub LoadMapData")

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando mapas."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempINT As Integer
Dim npcfile As String
Dim Tiempo As Long
Dim NoSirve As Boolean

On Error GoTo MAN


Dim Leer As New clsLeerInis

Tiempo = GetTickCount

'NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
frmCargando.Label1(0).Caption = "Preparando..."

Leer.Abrir DatPath & "Map.dat"
NumMaps = val(Leer.DarValor("INIT", "NumMaps"))
MapPath = Leer.DarValor("INIT", "MapPath")

frmCargando.Cargar.Min = 0
frmCargando.Cargar.max = NumMaps
frmCargando.Cargar.Value = 0
frmGeneral.ProG1.Min = 0
frmGeneral.ProG1.max = NumMaps
frmGeneral.ProG1.Value = 0
frmGeneral.ProG1.Visible = True
'MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

' [GS] Correccion para mapa pretoriano
If MAPA_PRETORIANO <> 0 Then
    'MsgBox "Activado Sistema Pretoriano, se cargaran todos los mapas hasta " & IIf(NumMaps < MAPA_PRETORIANO, MAPA_PRETORIANO, NumMaps) & "."
    If NumMaps < MAPA_PRETORIANO Then
        NumMaps = MAPA_PRETORIANO
        frmCargando.Cargar.max = NumMaps
        frmGeneral.ProG1.max = NumMaps
    End If
End If
' [/GS]

' [GS]
Call LogCOSAS("Mapas", "Cargado cantidad de mapas que es " & NumMaps, False)
' [/GS]


ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    DoEvents
    If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando mapas. " & str(Map) & " de " & str(NumMaps)

    NoSirve = val(Leer.DarValor("MAP" & Map, "NoCargar"))
    If NoSirve = False Then
        MapInfo(Map).Cargado = False
    
        Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        Seek #1, 1
        
        'inf
        Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        Seek #2, 1
        
         'map Header
        Get #1, , MapInfo(Map).MapVersion
        Get #1, , MiCabecera
        Get #1, , TempINT
        Get #1, , TempINT
        Get #1, , TempINT
        Get #1, , TempINT
    
        'inf Header
        Get #2, , TempINT
        Get #2, , TempINT
        Get #2, , TempINT
        Get #2, , TempINT
        Get #2, , TempINT
            
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                '.dat file
                Get #1, , MapData(Map, X, Y).Blocked
                
                For LoopC = 1 To 4
                    Get #1, , MapData(Map, X, Y).Graphic(LoopC)
                Next LoopC
                
                Get #1, , MapData(Map, X, Y).trigger
                Get #1, , TempINT
                
                                    
                '.inf file
                Get #2, , MapData(Map, X, Y).TileExit.Map
                Get #2, , MapData(Map, X, Y).TileExit.X
                Get #2, , MapData(Map, X, Y).TileExit.Y
                
                'Get and make NPC
                Get #2, , MapData(Map, X, Y).NpcIndex
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    
                    If MapData(Map, X, Y).NpcIndex > 499 Then
                            npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                            npcfile = DatPath & "NPCs.dat"
                    End If
                    
                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    End If
                    
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                    
                    Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If
    
                'Get and make Object
                Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
                Get #2, , MapData(Map, X, Y).OBJInfo.Amount
    
                'Space holder for future expansion (Objects, ect.
                Get #2, , DummyInt
                Get #2, , DummyInt
            
            Next X
        Next Y
    
       
        Close #1
        Close #2
    
      
        MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
        MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
        MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
        MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
        MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
        
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
            MapInfo(Map).Pk = True
        Else
            MapInfo(Map).Pk = False
        End If
        MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")
        MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
        MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
        MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))
        MapInfo(Map).MagiaSinEfecto = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MagiaSinEfecto"))
        
        ' Reset Moficaciones
        MapInfo(Map).Datos = False
        MapInfo(Map).Bloqueos = False
        MapInfo(Map).NPCs = False
        MapInfo(Map).Objs = False
        MapInfo(Map).Triggers = False
        MapInfo(Map).Telep = False
        
        MapInfo(Map).Cargado = True
        
    Else
        MapInfo(Map).Cargado = False
    End If
    
    If frmCargando.Visible Then
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
        frmCargando.Label1(0).Caption = "Mapa " & str(Map) & " de " & str(NumMaps)
        frmCargando.SetFocus
    End If
    If frmGeneral.Visible Then
        frmGeneral.ProG1.Value = frmGeneral.ProG1.Value + 1
        frmGeneral.Estado.SimpleText = "Cargando Mapas. " & str(Map) & " de " & str(NumMaps)
    End If

Next Map


Tiempo = GetTickCount - Tiempo
Call LogCOSAS("Tiempos", "Cargado de Mapas (" & NumMaps & ") " & str(Tiempo / 1000) & " segundos.", False)


frmGeneral.ProG1.Visible = False

Exit Sub

MAN:
    Call MsgBox("Error durante la carga de mapas (no del backup)." & vbCrLf & "El mapa " & Map & " contiene errores.", vbCritical)
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

    
End Sub

' [GS]
Public Sub LoadOpcsINI()

On Error Resume Next ' Para que no alla overflow err=3, pero no es aka?

' [GS]
' Especiales :D
HayTorneo = False   'No comienza con torneo :P
HayQuest = False    'No comienza con quest :P tambien :D
HayConsulta = False ' Menos comienza en consulta :P
QuienConsulta = 0
MapaDeTorneo = 0
PotsEnTorneo = True
MapaAventura = 0
MapaAgite = 0
Usando9999 = False
NoMensajeANW = False
MeditarChicoHasta = 15
MeditarMedioHasta = 30
' [/GS]

Call LeerConfigAntiChit
Call LeerConfigOpciones
Call LeerConfigClicks
Call LeerConfigMaximos
Call LeerConfigPrecios
Call LeerConfigPorcentajes
Call LeerConfigResucitar
Call LeerConfigUsuarios
Call LeerConfigExperiencia
Call LeerConfigMeditacion
Call LeerConfigAtributos
Call LeerConfigFacciones
Call LeerConfigNewbie
Call LeerConfigClanes
Call LeerConfigTorneo
Call LeerConfigAventura
Call LeerConfigCounter
Call LeerConfigServidor


Call LoadEstadisticas

End Sub
' [/GS]

' [GS]
Sub LoadEstadisticas()
On Error Resume Next
' SERVER
RecordUsuarios = val(GetVar(IniPath & "Estadisticas.ini", "SERVER", "RecordUsuarios"))
If RecordUsuarios < 0 Then RecordUsuarios = 0
If RecordUsuarios > tInt Then RecordUsuarios = tInt

' POWA
' Ponemos el mas powa base
ElMasPowa = GetVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Nombre")
If ElMasPowa = "" Then ElMasPowa = "nadie"
LvlDelPowa = val(GetVar(IniPath & "Estadisticas.ini", "POWA-LVL", "Level"))
If LvlDelPowa < 0 Then
    LvlDelPowa = 0
    ElMasPowa = "nadie"
End If
If LvlDelPowa > STAT_MAXELV Then
    LvlDelPowa = 0
    ElMasPowa = "nadie"
End If

PKNombre = GetVar(IniPath & "Estadisticas.ini", "POWA-PK", "Nombre")
If PKNombre = "" Then PKNombre = "nadie"
PKmato = val(GetVar(IniPath & "Estadisticas.ini", "POWA-PK", "Cantidad"))
If PKmato < 0 Then
    PKmato = 0
    PKNombre = "nadie"
End If
If PKmato > tLong Then
    PKmato = 0
    PKNombre = "nadie"
End If


MaxTINombre = GetVar(IniPath & "Estadisticas.ini", "POWA-TO", "Nombre")
If MaxTINombre = "" Then MaxTINombre = "nadie"
MaxTiempoOn = val(GetVar(IniPath & "Estadisticas.ini", "POWA-TO", "TiempoOnline"))
If MaxTiempoOn < 0 Then
    MaxTiempoOn = 0
    MaxTINombre = "nadie"
End If
If MaxTiempoOn > tLong Then
    MaxTiempoOn = 0
    MaxTINombre = "nadie"
End If

' LOTERIA
Pozo_Loteria = val(GetVar(IniPath & "Estadisticas.ini", "LOTERIA", "Pozo_Loteria"))
If Pozo_Loteria < 1000000 Then Pozo_Loteria = 1000000
If Pozo_Loteria > tLong Then Pozo_Loteria = tLong

End Sub
' [/GS]

Sub LoadSini()
On Error Resume Next ' Para que no alla overflow err=3?? no es aka?
Dim Temporal As Long
Dim Temporal1 As Long
Dim LoopC As Integer

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp") ' Hiper-AO
ServerName = GetVar(IniPath & "Server.ini", "INIT", "ServerName")

' [EL OSO] Pretorian Map
MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))
' [/EL OSO]
If MAPA_PRETORIANO <= 0 Then MAPA_PRETORIANO = 0


Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
' [GS] Puerto invalido
If Puerto < 1 Or Puerto > 65000 Then
    Call Alerta("Puerto invalido en Server.ini, [INIT] StartPort")
    Call Alerta("Corregido a Puerto 7666")
    Puerto = 7666
End If
' [/GS]

HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
ULTIMAVERSION2 = GetVar(IniPath & "Server.ini", "INIT", "Version1")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
'MsgBox "Sini 1"
ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))

ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))

'MsgBox "Sini 2"
ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))

If ClientsCommandsQueue <> 0 Then
        frmGeneral.CmdExec.Enabled = True
Else
        frmGeneral.CmdExec.Enabled = False
End If
'MsgBox "Sini 3"

'Start pos
StartPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
'FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
 
StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
'FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
 
SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
'FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
 
StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
'FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
 
IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
'FrmInterv.txtIntervaloSed.Text = IntervaloSed
 
IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
'FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
'FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
'FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
'FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
'FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
'FrmInterv.txtIntervaloWavFx.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
'FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
'FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
'FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmGeneral.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
'FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmGeneral.NpcAtaca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
'FrmInterv.txtNPCPuedeAtacar.Text = frmMain.NpcAtaca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
'FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
'FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmGeneral.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

frmGeneral.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
'FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs <> 0 Then If MinutosWs < 60 Then MinutosWs = 60

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
'Ressurect pos
ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))

'Max users
MaxUsers = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers < 1 Then MaxUsers = 10
ReDim UserList(1 To MaxUsers) As User

ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
For LoopC = 0 To UBound(MD5s)
    MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
    MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 53)
Next LoopC

' Ciudades!!!!
Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

' 0.12b3
Prision.Map = GetVar(DatPath & "Ciudades.dat", "Prision", "Mapa")
Prision.X = GetVar(DatPath & "Ciudades.dat", "Prision", "X")
Prision.Y = GetVar(DatPath & "Ciudades.dat", "Prision", "Y")
Libertad.Map = GetVar(DatPath & "Ciudades.dat", "Libertad", "Mapa")
Libertad.X = GetVar(DatPath & "Ciudades.dat", "Libertad", "X")
Libertad.Y = GetVar(DatPath & "Ciudades.dat", "Libertad", "Y")
' reparo bugs
If Prision.Map <= 0 Then
    Prision.Map = 66
End If
If Prision.X <= 0 Then
    Prision.X = 74
End If
If Prision.Y <= 0 Then
    Prision.Y = 47
End If
If Libertad.Map <= 0 Then
    Libertad.Map = 66
End If
If Libertad.X <= 0 Then
    Libertad.X = 75
End If
If Libertad.Y <= 0 Then
    Libertad.Y = 65
End If



If frmCargando.Visible = True Then Exit Sub

' [GS] Corrige un bug milenario
Call ReloadSokcet
Dim i As Integer
For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged = False Then Call CloseSocket(i)
Next i
' [/GS]

End Sub

Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(UserIndex As Integer, UserFile As String)
On Error GoTo errhandler

Dim OldUserHead As Long

If FileExist(UserFile, vbNormal) Then
       If UserList(UserIndex).flags.Muerto = 1 Then
        OldUserHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
       End If
       Kill UserFile
End If

Dim LoopC As Integer

'[Sicarul] ' Hiper-AO
Call WriteVar(UserFile, "FLAGS", "Casado", UserList(UserIndex).flags.Casado)
'[/Sicarul]
Call WriteVar(UserFile, "FLAGS", "Muerto", val(UserList(UserIndex).flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "Escondido", val(UserList(UserIndex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Hambre", val(UserList(UserIndex).flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", val(UserList(UserIndex).flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", val(UserList(UserIndex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(UserIndex).flags.ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(UserIndex).flags.Navegando))
' [GS]
Call WriteVar(UserFile, "FLAGS", "TiempoOnline", val(UserList(UserIndex).flags.TiempoOnline))
' [/GS]
Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(UserIndex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", val(UserList(UserIndex).flags.Paralizado))

Call WriteVar(UserFile, "COUNTERS", "Pena", val(UserList(UserIndex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", val(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", val(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", val(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", val(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", val(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", val(UserList(UserIndex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", val(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", val(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", val(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", val(UserList(UserIndex).Faccion.RecompensasReal))
' 0.12b1
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", val(UserList(UserIndex).Faccion.Reenlistadas))

Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(UserIndex).GuildInfo.EsGuildLeader))
Call WriteVar(UserFile, "GUILD", "Echadas", val(UserList(UserIndex).GuildInfo.Echadas))
Call WriteVar(UserFile, "GUILD", "Solicitudes", val(UserList(UserIndex).GuildInfo.Solicitudes))
Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(UserList(UserIndex).GuildInfo.SolicitudesRechazadas))
Call WriteVar(UserFile, "GUILD", "VecesFueGuildLeader", val(UserList(UserIndex).GuildInfo.VecesFueGuildLeader))
Call WriteVar(UserFile, "GUILD", "YaVoto", val(UserList(UserIndex).GuildInfo.YaVoto))
Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(UserIndex).GuildInfo.FundoClan))

Call WriteVar(UserFile, "GUILD", "GuildName", UserList(UserIndex).GuildInfo.GuildName)
Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(UserIndex).GuildInfo.ClanFundado)
Call WriteVar(UserFile, "GUILD", "ClanesParticipo", str(UserList(UserIndex).GuildInfo.ClanesParticipo))
Call WriteVar(UserFile, "GUILD", "GuildPts", str(UserList(UserIndex).GuildInfo.GuildPoints))

'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, val(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).Email)

Call WriteVar(UserFile, "INIT", "Genero", Num2Gen(UserList(UserIndex).genero))
Call WriteVar(UserFile, "INIT", "Raza", Num2Raza(UserList(UserIndex).raza))
Call WriteVar(UserFile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", Num2Clase(UserList(UserIndex).clase))
Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).Password)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).desc)

Call WriteVar(UserFile, "INIT", "Heading", str(UserList(UserIndex).Char.Heading))

Call WriteVar(UserFile, "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", str(UserList(UserIndex).Char.Body))
End If

Call WriteVar(UserFile, "INIT", "Arma", str(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", str(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", str(UserList(UserIndex).Char.CascoAnim))

' [GS] Aventura
Call WriteVar(UserFile, "AVENTURA", "IniPos", UserList(UserIndex).flags.AV_Lugar)
Call WriteVar(UserFile, "AVENTURA", "Tiempo", val(UserList(UserIndex).flags.AV_Tiempo))
' [/GS]

If Inbaneable(UserList(UserIndex).Name) Then UserList(UserIndex).IP = "No Guardado!"

Call WriteVar(UserFile, "INIT", "LastIP", UserList(UserIndex).IP)
Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)

Call WriteVar(UserFile, "STATS", "GLD", str(UserList(UserIndex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", str(UserList(UserIndex).Stats.banco))

Call WriteVar(UserFile, "STATS", "MET", str(UserList(UserIndex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", str(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", str(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "FIT", str(UserList(UserIndex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", str(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", str(UserList(UserIndex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", str(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", str(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", str(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", str(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", str(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", str(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", str(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", str(UserList(UserIndex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", str(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", str(UserList(UserIndex).Stats.exp))
Call WriteVar(UserFile, "STATS", "ELV", str(UserList(UserIndex).Stats.ELV))
Call WriteVar(UserFile, "STATS", "ELU", str(UserList(UserIndex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", val(UserList(UserIndex).Stats.UsuariosMatados))
Call WriteVar(UserFile, "MUERTES", "CrimMuertes", val(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", val(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", CStr(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", CStr(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", CStr(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", CStr(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", CStr(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", CStr(UserList(UserIndex).Invent.MunicionEqpSlot))
Call WriteVar(UserFile, "Inventory", "HerramientaSlot", CStr(UserList(UserIndex).Invent.HerramientaEqpSlot))
Call WriteVar(UserFile, "Inventory", "Accesorio1Slot", CStr(UserList(UserIndex).Invent.Accesorio1EqpSlot))
Call WriteVar(UserFile, "Inventory", "Accesorio2Slot", CStr(UserList(UserIndex).Invent.Accesorio2EqpSlot))

'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", val(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", val(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", val(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", val(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", val(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", val(UserList(UserIndex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Call WriteVar(UserFile, "REP", "Promedio", val(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next

Dim NroMascotas As Long
NroMascotas = UserList(UserIndex).NroMacotas

For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", str(NroMascotas))

Call WriteVar(UserFile, "ADMINISTRACION", "EnPrueba", IIf(UserList(UserIndex).Administracion.EnPrueba = True, 1, 0))

Call WriteVar(UserFile, "ADMINISTRACION", "Activado", IIf(UserList(UserIndex).Administracion.Activado = True, 1, 0))

Call WriteVar(UserFile, "ADMINISTRACION", "Config", UserList(UserIndex).Administracion.Config)

Call WriteVar(UserFile, "ADMINISTRACION", "CP", val(UserList(UserIndex).Administracion.MaxCP))
For LoopC = 1 To UserList(UserIndex).Administracion.MaxCP
    Call WriteVar(UserFile, "ADMINISTRACION", "CP" & LoopC, UserList(UserIndex).Administracion.CP(LoopC))
Next
'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).Char.Head = iCabezaMuerto
End If

' 0.12b1
Call WriteVar(UserFile, "CONSEJO", "PERTENECE", val(IIf(UserList(UserIndex).flags.PertAlCons = True, 1, 0)))
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", val(IIf(UserList(UserIndex).flags.PertAlConsCaos = True, 1, 0)))


Exit Sub

errhandler:
Call LogError("Error en SaveUser: " & Err.Number & " " & Err.Description)

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Criminal = (L < 0)

End Function




Sub BackUPnPc(NpcIndex As Integer)
On Error GoTo ErroRR
'Call LogTarea("Sub BackUPnPc NpcIndex:" & NpcIndex)
Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer

NpcNumero = Npclist(NpcIndex).Numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
' [GS] Intercambia
Call WriteVar(npcfile, "NPC" & NpcNumero, "Intercambia", val(Npclist(NpcIndex).Intercambia))
' [/GS]
' [GS] Combate
Call WriteVar(npcfile, "NPC" & NpcNumero, "Combate", val(Npclist(NpcIndex).Combate))
' [/GS]
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))

'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Comercia = 1 Then
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
        Else
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "")
        End If
   Next
End If

Exit Sub
ErroRR:
Call LogError("Error en BackUPnPc: " & Err.Number & " " & Err.Description)
Call RepararDAT("bkNPCs.dat", "bkNPCs_2.dat")
End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
On Error GoTo errorx
'Call LogTarea("Sub CargarNpcBackUp NpcIndex:" & NpcIndex & " NpcNumber:" & NpcNumber)
Dim ParteError As Integer
ParteError = 0
'Status
If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Cargando backup Npc"



Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "bkNPCs.dat"
End If

ParteError = 1

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

ParteError = 2

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

ParteError = 3

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Intercambia = val(GetVar(npcfile, "NPC" & NpcNumber, "Intercambia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) ' * 8 'Lo multiplique!!!:P

ParteError = 4

Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD")) ' * 8 'Esto tambien lo multiplique!!!:P
If Npclist(NpcIndex).GiveGLD < 1 Then Npclist(NpcIndex).GiveGLD = 0
Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

ParteError = 5

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))

ParteError = 6

Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If

ParteError = 7

Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

ParteError = 8

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

ParteError = 9

'Tipo de items con los que comercia
If Npclist(NpcIndex).Comercia <> 0 Then Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

ParteError = 10


ParteError = 11
Call HacerDAT("bkNPCs.dat")
ParteError = 12

Exit Sub
errorx:
If ParteError > 10 Then Exit Sub ' Sale bien igual carajo!
Call MsgBox("Error durante la carga de NPC de backup." & vbCrLf & "El NPC " & NpcNumber & " tiene un error.", vbCritical)
Call LogError("ERROR Cargando NPC de Backup - NPC: " & NpcNumber & " INDEX: " & NpcIndex & " ParteError: " & ParteError)
If FileExist(App.Path & "\Dat\bkNPCs.dat", vbArchive) Then
    BorrarArchivo (App.Path & "\Dat\bkNPCs.dat")
    If FileExist(App.Path & "\Dat\Seguridad\bkNPCs.dat", vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\bkNPCs.dat", App.Path & "\Dat\bkNPCs.dat")
    ElseIf FileExist(App.Path & "\Dat\Seguridad\bkNPCs_2.dat", vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\bkNPCs_2.dat", App.Path & "\Dat\bkNPCs.dat")
    ElseIf FileExist(App.Path & "\Dat\Seguridad\NPCs.dat", vbArchive) Then
        Call FileCopy(App.Path & "\Dat\NPCs.dat", App.Path & "\Dat\bkNPCs.dat")
    End If
    If FileExist(App.Path & "\Dat\bkNPCs.dat", vbArchive) Then
        MsgBox "BackupNPCs reparado. Vuelva a ejecutar el Servidor.", vbInformation
        End
    Else
        MsgBox "No se pudo reparar bkNPCs.bat, de la carpeta DAT" & vbCrLf & "Intente repararlo manualmente, con una copia anterior.", vbCritical
        End
    End If
End If
End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub




Function FileSize(ByVal FileNamE As String) As Long
On Error GoTo FalloFile
Dim nFileNum As Integer
Dim lFileSize As Long
nFileNum = FreeFile
FileSize = -1
Open FileNamE For Input As nFileNum
lFileSize = LOF(nFileNum)
Close nFileNum
FileSize = lFileSize

Exit Function
FalloFile:
FileSize = -1
End Function

Public Sub VerificarEstenTodosLosArchivosYCarpetas()
On Error Resume Next
Dim Patho As String
Dim i As Long
'\Dat
DoEvents
If FileExist(App.Path & "\Opciones.ini", vbArchive) = False Then
    If FileExist(App.Path & "\Opciones.ini.default", vbArchive) = False Then
        MsgBox "ALERTA: El servidor no se encuentra configurado. Falta: opciones.ini", vbCritical
    Else
        Call FileCopy(App.Path & "\Opciones.ini.default", App.Path & "\Opciones.ini")
    End If
End If
If FileExist(App.Path & "\Server.ini", vbArchive) = False Then
    MsgBox "ALERTA: El servidor no se encuentra configurado. Falta: server.ini", vbCritical
End If
If FileExist(App.Path & "\Estadisticas.ini", vbArchive) = False Then
    If FileExist(App.Path & "\Estadisticas.ini.default", vbArchive) = False Then
        MsgBox "ALERTA: No se encuentran las estadisticas del servidor. Falta: estadisticas.ini", vbCritical
    Else
        Call FileCopy(App.Path & "\Estadisticas.ini.default", App.Path & "\Estadisticas.ini")
    End If
End If
DoEvents

If val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "NOHacerDiagnostico")) = 1 Then Exit Sub

If FileExist(App.Path & "\Dat\", vbDirectory) Then
    ' Existe la carpeta
    Patho = App.Path & "\Dat\"
    If FileExist(Patho & "Seguridad\", vbDirectory) = False Then
        Call MkDir(Patho & "Seguridad\")
    End If
    If FileExist(Patho & "ArmadurasHerrero.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo ArmadurasHerrero.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\ArmadurasHerrero.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\ArmadurasHerrero.dat", Patho & "ArmadurasHerrero.dat")
            MsgBox "ArmadurasHerrero.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo ArmadurasHerrero.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "ArmadurasHerrero.dat") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(Patho & "Seguridad\ArmadurasHerrero.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\ArmadurasHerrero.dat")
            End If
            Call FileCopy(Patho & "ArmadurasHerrero.dat", Patho & "Seguridad\ArmadurasHerrero.dat") ' Copio el nuevo
        End If
    End If
    If FileExist(Patho & "ArmasHerrero.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo ArmasHerrero.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\ArmasHerrero.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\ArmasHerrero.dat", Patho & "ArmasHerrero.dat")
            MsgBox "ArmasHerrero.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo ArmasHerrero.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "ArmasHerrero.dat") > 8 Then ' Si tiene mas de 9 bytes
            If FileExist(Patho & "Seguridad\ArmasHerrero.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\ArmasHerrero.dat")
            End If
            Call FileCopy(Patho & "ArmasHerrero.dat", Patho & "Seguridad\ArmasHerrero.dat") ' Copio el nuevo
        End If
    End If
    If FileExist(Patho & "NPCs.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo NPCs.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\NPCs.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\NPCs.dat", Patho & "NPCs.dat")
            MsgBox "NPCs.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo NPCs.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "NPCs.dat") > 19999 Then ' Si tiene mas de 20 Kbytes
            If FileExist(Patho & "Seguridad\NPCs.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\NPCs.dat")
            End If
            Call FileCopy(Patho & "NPCs.dat", Patho & "Seguridad\NPCs.dat") ' Copio el nuevo
        End If
    End If
    If FileExist(Patho & "NPCs-HOSTILES.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo NPCs-HOSTILES.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\NPCs-HOSTILES.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\NPCs-HOSTILES.dat", Patho & "NPCs-HOSTILES.dat")
            MsgBox "NPCs-HOSTILES.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo NPCs-HOSTILES.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "NPCs-HOSTILES.dat") > 10999 Then ' Si tiene mas de 11 Kbytes
            If FileExist(Patho & "Seguridad\NPCs-HOSTILES.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\NPCs-HOSTILES.dat")
            End If
            Call FileCopy(Patho & "NPCs-HOSTILES.dat", Patho & "Seguridad\NPCs-HOSTILES.dat") ' Copio el nuevo
        End If
    End If
    If FileExist(Patho & "bkNPCs.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo bkNPCs.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\bkNPCs.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\bkNPCs.dat", Patho & "bkNPCs.dat")
            MsgBox "NPCs-HOSTILES.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo bkNPCs.dat, de la Carpeta DAT." & vbCrLf & "Se intentara reparar remplazandolo por NPCs.dat", vbInformation
            Call FileCopy(Patho & "NPCs.dat", Patho & "bkNPCs.dat")
            DoEvents
            If FileExist(Patho & "bkNPCs.dat", vbArchive) = False Then
                MsgBox "No se pudo reparar bkNPCs.dat, en la Carpeta DAT", vbCritical
                End
            End If
        End If
    Else
        If FileSize(Patho & "bkNPCs.dat") > 19999 Then ' Si tiene mas de 20 Kbytes
            If FileExist(Patho & "Seguridad\bkNPCs_2.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\bkNPCs_2.dat")
            End If
            Call FileCopy(Patho & "bkNPCs.dat", Patho & "Seguridad\bkNPCs_2.dat") ' Copio el nuevo
        End If
    End If
'    If FileExist(Patho & "Body.dat", vbArchive) = False Then
'        MsgBox "No se encuentra el archivo Body.dat, en la Carpeta DAT", vbInformation
'        If FileExist(Patho & "Seguridad\Body.dat", vbArchive) Then
'            Call FileCopy(Patho & "Seguridad\Body.dat", Patho & "Body.dat")
'            MsgBox "Body.dat reparado, desde la copia de seguridad."
'        Else
'            MsgBox "No existen copias de seguridad del archivo Body.dat, de la Carpeta DAT.", vbCritical
'            End
'        End If
'    Else
'        If FileSize(Patho & "Body.dat") > 1999 Then ' Si tiene mas de 2 Kbytes
'            If FileExist(Patho & "Seguridad\Body.dat", vbArchive) Then ' Borro el de seguridad existente
'                Call BorrarArchivo(Patho & "Seguridad\Body.dat")
'            End If
'            Call FileCopy(Patho & "Body.dat", Patho & "Seguridad\Body.dat") ' Copio el nuevo
'        End If
'    End If
    DoEvents
    If FileExist(Patho & "Ciudades.Dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Ciudades.Dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Ciudades.Dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Ciudades.Dat", Patho & "Ciudades.Dat")
            MsgBox "Ciudades.Dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Ciudades.Dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Ciudades.Dat") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(Patho & "Seguridad\Ciudades.Dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Ciudades.Dat")
            End If
            Call FileCopy(Patho & "Ciudades.Dat", Patho & "Seguridad\Ciudades.Dat") ' Copio el nuevo
        End If
    End If
    
'    If FileExist(Patho & "Head.dat", vbArchive) = False Then
'        MsgBox "No se encuentra el archivo Head.dat, en la Carpeta DAT", vbInformation
'        If FileExist(Patho & "Seguridad\Head.dat", vbArchive) Then
'            Call FileCopy(Patho & "Seguridad\Head.dat", Patho & "Head.dat")
'            MsgBox "Head.dat reparado, desde la copia de seguridad."
'        Else
'            MsgBox "No existen copias de seguridad del archivo Head.dat, de la Carpeta DAT.", vbCritical
'            End
'        End If
'    Else
'        If FileSize(Patho & "Head.dat") > 19 Then ' Si tiene mas de 20 bytes
'            If FileExist(Patho & "Seguridad\Head.dat", vbArchive) Then ' Borro el de seguridad existente
'                Call BorrarArchivo(Patho & "Seguridad\Head.dat")
'            End If
'            Call FileCopy(Patho & "Head.dat", Patho & "Seguridad\Head.dat") ' Copio el nuevo
'        End If
'    End If
    
    If FileExist(Patho & "Hechizos.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Hechizos.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Hechizos.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Hechizos.dat", Patho & "Hechizos.dat")
            MsgBox "Hechizos.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Hechizos.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Hechizos.dat") > 11999 Then ' Si tiene mas de 12 Kbytes
            If FileExist(Patho & "Seguridad\Hechizos_2.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Hechizos_2.dat")
            End If
            Call FileCopy(Patho & "Hechizos.dat", Patho & "Seguridad\Hechizos_2.dat") ' Copio el nuevo
        End If
    End If
    DoEvents
    If FileExist(Patho & "Help.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Help.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Help.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Help.dat", Patho & "Help.dat")
            MsgBox "Help.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Help.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Help.dat") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(Patho & "Seguridad\Help.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Help.dat")
            End If
            Call FileCopy(Patho & "Help.dat", Patho & "Seguridad\Help.dat") ' Copio el nuevo
        End If
    End If
    
    If FileExist(Patho & "Invokar.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Invokar.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Invokar.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Invokar.dat", Patho & "Invokar.dat")
            MsgBox "Invokar.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Invokar.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Invokar.dat") > 199 Then ' Si tiene mas de 200 bytes
            If FileExist(Patho & "Seguridad\Invokar.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Invokar.dat")
            End If
            Call FileCopy(Patho & "Invokar.dat", Patho & "Seguridad\Invokar.dat") ' Copio el nuevo
        End If
    End If
    
    If FileExist(Patho & "Map.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Map.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Map.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Map.dat", Patho & "Map.dat")
            MsgBox "Map.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Map.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Map.dat") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(Patho & "Seguridad\Map.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Map.dat")
            End If
            Call FileCopy(Patho & "Map.dat", Patho & "Seguridad\Map.dat") ' Copio el nuevo
        End If
    End If
    
    If FileExist(Patho & "Motd.ini", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Motd.ini, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Motd.ini", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Motd.ini", Patho & "Motd.ini")
            MsgBox "Motd.ini reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Motd.ini, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Motd.ini") > 9 Then ' Si tiene mas de 10 bytes
            If FileExist(Patho & "Seguridad\Motd.ini", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Motd.ini")
            End If
            Call FileCopy(Patho & "Motd.ini", Patho & "Seguridad\Motd.ini") ' Copio el nuevo
        End If
    End If
    DoEvents
    If FileExist(Patho & "NombresInvalidos.txt", vbArchive) = False Then
        MsgBox "No se encuentra el archivo NombresInvalidos.txt, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\NombresInvalidos.txt", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\NombresInvalidos.txt", Patho & "NombresInvalidos.txt")
            MsgBox "NombresInvalidos.txt reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo NombresInvalidos.txt, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "NombresInvalidos.txt") > 9 Then ' Si tiene mas de 9 bytes
            If FileExist(Patho & "Seguridad\NombresInvalidos.txt", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\NombresInvalidos.txt")
            End If
            Call FileCopy(Patho & "NombresInvalidos.txt", Patho & "Seguridad\NombresInvalidos.txt") ' Copio el nuevo
        End If
    End If
    
    If FileExist(Patho & "Obj.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Obj.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\Obj.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\Obj.dat", Patho & "Obj.dat")
            MsgBox "Obj.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo Obj.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "Obj.dat") > 109999 Then ' Si tiene mas de 110 Kbytes
            If FileExist(Patho & "Seguridad\Obj_2.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\Obj_2.dat")
            End If
            Call FileCopy(Patho & "Obj.dat", Patho & "Seguridad\Obj_2.dat") ' Copio el nuevo
        End If
    End If
    
    If FileExist(Patho & "ObjCarpintero.dat", vbArchive) = False Then
        MsgBox "No se encuentra el archivo Obj.dat, en la Carpeta DAT", vbInformation
        If FileExist(Patho & "Seguridad\ObjCarpintero.dat", vbArchive) Then
            Call FileCopy(Patho & "Seguridad\ObjCarpintero.dat", Patho & "ObjCarpintero.dat")
            MsgBox "ObjCarpintero.dat reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo ObjCarpintero.dat, de la Carpeta DAT.", vbCritical
            End
        End If
    Else
        If FileSize(Patho & "ObjCarpintero.dat") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(Patho & "Seguridad\ObjCarpintero.dat", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(Patho & "Seguridad\ObjCarpintero.dat")
            End If
            Call FileCopy(Patho & "ObjCarpintero.dat", Patho & "Seguridad\ObjCarpintero.dat") ' Copio el nuevo
        End If
    End If
    
Else
    MsgBox "No se encuentra el directorio DAT!" & vbCrLf & "Por favor, copie todo el material que esta carpeta necesita.", vbCritical
    Call MkDir(App.Path & "\Dat\")
    DoEvents
    End
End If
DoEvents
If FileExist(App.Path & "\Bugs\", vbDirectory) = False Then
    'MsgBox "La carpeta Bugs no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Bugs\")
    If FileExist(App.Path & "\Bugs\", vbDirectory) = True Then
        'MsgBox "Carpeta Bugs reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Bugs, creela manualmente.", vbCritical
        End
    End If
End If

If FileExist(App.Path & "\Charfile\", vbDirectory) = False Then
    'MsgBox "La carpeta Charfile no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Charfile\")
    If FileExist(App.Path & "\Charfile\", vbDirectory) = True Then
        'MsgBox "Carpeta Charfile reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Charfile, creela manualmente.", vbCritical
        End
    End If
End If

If FileExist(App.Path & "\chrbackup\", vbDirectory) = False Then
    'MsgBox "La carpeta chrbackup no se encuentra.", vbInformation
    Call MkDir(App.Path & "\chrbackup\")
    If FileExist(App.Path & "\chrbackup\", vbDirectory) = True Then
        'MsgBox "Carpeta chrbackup reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta chrbackup, creela manualmente.", vbCritical
        End
    End If
End If
DoEvents
If FileExist(App.Path & "\Foros\", vbDirectory) = False Then
    'MsgBox "La carpeta Foros no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Foros\")
    If FileExist(App.Path & "\Foros\", vbDirectory) = True Then
        'MsgBox "Carpeta Foros reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Foros, creela manualmente.", vbCritical
        End
    End If
End If

If FileExist(App.Path & "\Guilds\", vbDirectory) = False Then
    MsgBox "La carpeta Guilds no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Guilds\")
    If FileExist(App.Path & "\Guilds\", vbDirectory) = True Then
        MsgBox "Carpeta Guilds reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Guilds, creela manualmente y copie su contenido normal.", vbCritical
        End
    End If
Else
    If FileExist(App.Path & "\Guilds\Seguridad\", vbDirectory) = False Then
        Call MkDir(App.Path & "\Guilds\Seguridad\")
    End If
    If FileExist(App.Path & "\Guilds\GuildsInfo.inf", vbArchive) = False Then
        MsgBox "No se encuentra GuildsInfo.inf, en la Carpeta Guilds", vbInformation
        If FileExist(App.Path & "\Guilds\Seguridad\GuildsInfo.inf", vbArchive) Then
            Call FileCopy(App.Path & "\Guilds\Seguridad\GuildsInfo.inf", App.Path & "\Guilds\GuildsInfo.inf")
            MsgBox "GuildsInfo.inf reparado, desde la copia de seguridad."
        Else
            MsgBox "No existen copias de seguridad del archivo GuildsInfo.inf, de la Carpeta Guilds.", vbCritical
            End
        End If
    Else
        If FileSize(App.Path & "\Guilds\GuildsInfo.inf") > 19 Then ' Si tiene mas de 20 bytes
            If FileExist(App.Path & "\Guilds\Seguridad\GuildsInfo.inf", vbArchive) Then ' Borro el de seguridad existente
                Call BorrarArchivo(App.Path & "\Guilds\Seguridad\GuildsInfo.inf")
            End If
            Call FileCopy(App.Path & "\Guilds\GuildsInfo.inf", App.Path & "\Guilds\Seguridad\GuildsInfo.inf") ' Copio el nuevo
        End If
    End If
End If

If FileExist(App.Path & "\Logs\", vbDirectory) = False Then
    'MsgBox "La carpeta Logs no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Logs\")
    If FileExist(App.Path & "\Logs\", vbDirectory) = True Then
        'MsgBox "Carpeta Logs reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Logs, creela manualmente.", vbCritical
        End
    End If
End If

If FileExist(App.Path & "\Logs\Consejeros\", vbDirectory) = False Then
    Call MkDir(App.Path & "\Logs\Consejeros\")
End If

If FileExist(App.Path & "\Logs\Usuarios\", vbDirectory) = False Then
    Call MkDir(App.Path & "\Logs\Usuarios\")
End If

If FileExist(App.Path & "\Wav\", vbDirectory) = False Then
    'MsgBox "La carpeta Wav no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Wav\")
    If FileExist(App.Path & "\Wav\", vbDirectory) = True Then
        'MsgBox "Carpeta Wav reparada.", vbInformation
    Else
        MsgBox "Error al crear la carpeta Wav, creela manualmente.", vbCritical
        End
    End If
Else
    If FileExist(App.Path & "\Wav\Harp3.wav", vbArchive) = False Then
        Call Alerta("No se encuentra Harp3.wav, en la carpeta Wav")
        'MsgBox "El servidor continuara con la carga pero sin este archivo.", vbInformation
    End If
End If
DoEvents
Call VerificarMapas
DoEvents
End Sub


Sub VerificarMapas()
On Error GoTo FalloMap
Dim i As Integer

If FileExist(App.Path & "\Maps\", vbDirectory) = False Then
    MsgBox "La carpeta Maps no se encuentra.", vbInformation
    Call MkDir(App.Path & "\Maps\")
    If FileExist(App.Path & "\Maps\", vbDirectory) = True Then
        MsgBox "Carpeta Maps creada, pero faltan todos los mapas, copielos manualmente.", vbInformation
        End
    Else
        MsgBox "Error al crear la carpeta Maps, creela manualmente y copie todos los mapas en ella.", vbCritical
        End
    End If
Else
    NumMaps = val(GetVar(App.Path & "\Dat\Map.dat", "INIT", "NumMaps"))
    For i = 1 To NumMaps
        If i = 81 Or i = 82 Or i = 83 Or i = 84 Or i = 85 Or i = 117 Or i = 118 Or i = 119 Or i = 160 Or i = 161 Or i = 165 Then GoTo SiguienTeMapZ
        ' Si es 81, 82, 83, 84, 84, 117, 118 o 119 no son mapas validos
            ' Esta el mapa.dat, .inf y .map
            If FileSize(App.Path & "\Maps\mapa" & i & ".map") < 130000 Then
                MsgBox "El archivo mapa" & i & ".map se encuentra dañado. Remplazelo manualmente en la carpeta Maps.", vbCritical
                End
            End If
            If FileSize(App.Path & "\Maps\mapa" & i & ".inf") < 150000 Then
                MsgBox "El archivo mapa" & i & ".inf se encuentra dañado. Remplazelo manualmente en la carpeta Maps.", vbCritical
                End
            End If
            If FileSize(App.Path & "\Maps\mapa" & i & ".dat") < 10 Then
                MsgBox "El archivo mapa" & i & ".dat se encuentra dañado. Remplazelo manualmente en la carpeta Maps.", vbCritical
                End
            End If
SiguienTeMapZ:
    Next
End If

' WORLDBACKUP ' WORLDBACKUP
' WORLDBACKUP ' WORLDBACKUP
' WORLDBACKUP ' WORLDBACKUP

DoEvents
If FileExist(App.Path & "\WorldBackUp\", vbDirectory) = False Then
    'MsgBox "La carpeta WorldBackUp no se encuentra.", vbInformation
    Call MkDir(App.Path & "\WorldBackUp\")
    If FileExist(App.Path & "\WorldBackUp\", vbDirectory) = True Then
        'MsgBox "Carpeta WorldBackUp creada, si tiene una copia del Backup copiela manualmente.", vbInformation
        Exit Sub
    Else
        MsgBox "Error al crear la carpeta WorldBackUp, creela manualmente y si tiene una copia de su ultimo Backup funcional copiela.", vbCritical
        End
        Exit Sub
    End If
Else
    If FileExist(App.Path & "\WorldBackUp\Seguridad\", vbDirectory) = False Then
        Call MkDir(App.Path & "\WorldBackUp\Seguridad\")
    End If
    NumMaps = val(GetVar(App.Path & "\Dat\Map.dat", "INIT", "NumMaps"))
    For i = 1 To NumMaps
            ' Esta el mapa.dat, .inf y .map
            ' ARCHIVO MAP
        If FileExist(App.Path & "\WorldBackUp\map" & i & ".map", vbArchive) = True Then
            If FileSize(App.Path & "\WorldBackUp\map" & i & ".map") < 130000 Then
                MsgBox "El archivo map" & i & ".map se encuentra dañado, de la carpeta WorldBackUp.", vbInformation
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".map", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\map" & i & ".map")
                    Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & i & ".map", App.Path & "\WorldBackUp\map" & i & ".map")
                    If FileExist(App.Path & "\WorldBackUp\map" & i & ".map", vbArchive) = True Then
                        MsgBox "map" & i & ".map fue reparado con exito", vbInformation
                    Else
                        MsgBox "No se pudo reparar map" & i & ".map" & vbCrLf & "Hagalo manualmente o eliminelo.", vbCritical
                    End If
                Else
                    MsgBox "No existen copias de seguridad del mapa." & vbCrLf & "Tendra que reparar o eliminar manualmente map" & i & ".map, de la carpeta WorldBackUp.", vbCritical
                    End
                    Exit Sub
                End If
            Else
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".map", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\Seguridad\map" & i & ".map")
                End If
                Call FileCopy(App.Path & "\WorldBackUp\map" & i & ".map", App.Path & "\WorldBackUp\Seguridad\map" & i & ".map")
            End If
        End If
            ' ARCHIVO INF
        If FileExist(App.Path & "\WorldBackUp\map" & i & ".inf", vbArchive) = True Then
            If FileSize(App.Path & "\WorldBackUp\map" & i & ".inf") < 150000 Then
                MsgBox "El archivo map" & i & ".inf se encuentra dañado, de la carpeta WorldBackUp.", vbInformation
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".inf", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\map" & i & ".inf")
                    Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & i & ".inf", App.Path & "\WorldBackUp\map" & i & ".inf")
                    If FileExist(App.Path & "\WorldBackUp\map" & i & ".inf", vbArchive) = True Then
                        MsgBox "map" & i & ".inf fue reparado con exito", vbInformation
                    Else
                        MsgBox "No se pudo reparar map" & i & ".inf" & vbCrLf & "Hagalo manualmente o eliminelo.", vbCritical
                    End If
                Else
                    MsgBox "No existen copias de seguridad del mapa." & vbCrLf & "Tendra que reparar o eliminar manualmente map" & i & ".inf, de la carpeta WorldBackUp.", vbCritical
                    End
                    Exit Sub
                End If
            Else
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".inf", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\Seguridad\map" & i & ".inf")
                End If
                Call FileCopy(App.Path & "\WorldBackUp\map" & i & ".inf", App.Path & "\WorldBackUp\Seguridad\map" & i & ".inf")
            End If
        End If
        
        If FileExist(App.Path & "\WorldBackUp\map" & i & ".dat", vbArchive) = True Then
            ' ARCHIVO DAT
            If FileSize(App.Path & "\WorldBackUp\map" & i & ".dat") < 40 Then
                MsgBox "El archivo map" & i & ".dat se encuentra dañado, de la carpeta WorldBackUp.", vbInformation
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".dat", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\map" & i & ".dat")
                    Call FileCopy(App.Path & "\WorldBackUp\Seguridad\map" & i & ".dat", App.Path & "\WorldBackUp\map" & i & ".dat")
                    If FileExist(App.Path & "\WorldBackUp\map" & i & ".dat", vbArchive) = True Then
                        MsgBox "map" & i & ".dat fue reparado con exito", vbInformation
                    Else
                        MsgBox "No se pudo reparar map" & i & ".dat" & vbCrLf & "Hagalo manualmente o eliminelo.", vbCritical
                    End If
                Else
                    MsgBox "No existen copias de seguridad del mapa." & vbCrLf & "Tendra que reparar o eliminar manualmente map" & i & ".dat, de la carpeta WorldBackUp.", vbCritical
                    End
                    Exit Sub
                End If
            Else
                If FileExist(App.Path & "\WorldBackUp\Seguridad\map" & i & ".dat", vbArchive) = True Then
                    Call BorrarArchivo(App.Path & "\WorldBackUp\Seguridad\map" & i & ".dat")
                End If
                Call FileCopy(App.Path & "\WorldBackUp\map" & i & ".dat", App.Path & "\WorldBackUp\Seguridad\map" & i & ".dat")
            End If
        End If
    Next
End If

Exit Sub

FalloMap:

Call LogError("Error durante la Verificacion de Errores en Mapas - " & Err.Number & ":" & Err.Description)
Resume Next

End Sub

Sub HacerDAT(ByVal Original As String)
On Error Resume Next
If FileExist(App.Path & "\Dat\Seguridad\", vbDirectory) = False Then
    Call MkDir(App.Path & "\Dat\Seguridad\")
End If
If FileExist(App.Path & "\Dat\Seguridad\" & Original, vbArchive) = True Then ' Borro el de seguridad existente
    If FileSize(App.Path & "\Dat\Seguridad\" & Original) <> FileSize(App.Path & "\Dat\" & Original) Then
        Call BorrarArchivo(App.Path & "\Dat\Seguridad\" & Original)
        Call FileCopy(App.Path & "\Dat\" & Original, App.Path & "\Dat\Seguridad\" & Original) ' Copio el nuevo
    End If
Else
    If FileExist(App.Path & "\Dat\" & Original, vbArchive) = True Then
        If FileExist(App.Path & "\Dat\Seguridad\", vbDirectory) Then
            Call FileCopy(App.Path & "\Dat\" & Original, App.Path & "\Dat\Seguridad\" & Original) ' Copio el nuevo
        End If
    End If
End If
End Sub

Sub RepararDAT(ByVal Original As String, ByVal Segundo As String)
On Error Resume Next
Dim h As Long
If FileExist(App.Path & "\Dat\" & Original, vbArchive) Then
    BorrarArchivo (App.Path & "\Dat\" & Original)
    If FileExist(App.Path & "\Dat\Seguridad\" & Original, vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\" & Original, App.Path & "\Dat\" & Original)
    ElseIf FileExist(App.Path & "\Dat\Seguridad\" & Segundo, vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\" & Segundo, App.Path & "\Dat\" & Original)
    End If
    If FileExist(App.Path & "\Dat\" & Original, vbArchive) Then
        MsgBox Original & " reparados. Vuelva a ejecutar el Servidor.", vbInformation
        ' [GS] Hay alguien?
        frmGeneral.mnuCerrar.Checked = True
        For h = 1 To LastUser
            If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
                Call SendData(ToAll, 0, 0, "||<Host> A ocurrido un error..." & FONTTYPE_TALK & ENDC)
                Call SendData(ToAll, 0, 0, "||<Host> Cerramos, en 5 min abrimos..." & FONTTYPE_TALK & ENDC)
                Call GuardarUsuarios
                Call SaveGuildsDB
                End
                Exit Sub
            End If
        Next h
        Call SaveGuildsDB
        End
        ' [/GS]
    Else
        MsgBox "No se pudo reparar " & Original & ", de la carpeta DAT" & vbCrLf & "Intente repararlo manualmente, con una copia anterior.", vbCritical
        ' [GS] Hay alguien?
        frmGeneral.mnuCerrar.Checked = True
        For h = 1 To LastUser
            If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
                Call SendData(ToAll, 0, 0, "||<Host> A ocurrido un error..." & FONTTYPE_TALK & ENDC)
                Call SendData(ToAll, 0, 0, "||<Host> Cerramos, en 5 min abrimos..." & FONTTYPE_TALK & ENDC)
                Call GuardarUsuarios
                Call SaveGuildsDB
                End
                Exit Sub
            End If
        Next h
        Call SaveGuildsDB
        End
        ' [/GS]
    End If
Else
    If FileExist(App.Path & "\Dat\Seguridad\" & Original, vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\" & Original, App.Path & "\Dat\" & Original)
    ElseIf FileExist(App.Path & "\Dat\Seguridad\" & Segundo, vbArchive) Then
        Call FileCopy(App.Path & "\Dat\Seguridad\" & Segundo, App.Path & "\Dat\" & Original)
    End If
    If FileExist(App.Path & "\Dat\" & Original, vbArchive) Then
        MsgBox Original & " reparados. Vuelva a ejecutar el Servidor.", vbInformation
        ' [GS] Hay alguien?
        frmGeneral.mnuCerrar.Checked = True
        For h = 1 To LastUser
            If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
                Call SendData(ToAll, 0, 0, "||<Host> A ocurrido un error..." & FONTTYPE_TALK & ENDC)
                Call SendData(ToAll, 0, 0, "||<Host> Cerramos, en 5 min abrimos..." & FONTTYPE_TALK & ENDC)
                Call GuardarUsuarios
                Call SaveGuildsDB
                End
                Exit Sub
            End If
        Next h
        Call SaveGuildsDB
        End
        ' [/GS]
    Else
        MsgBox "No se pudo reparar " & Original & ", de la carpeta DAT" & vbCrLf & "Intente repararlo manualmente, con una copia anterior.", vbCritical
        ' [GS] Hay alguien?
        frmGeneral.mnuCerrar.Checked = True
        For h = 1 To LastUser
            If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
                Call SendData(ToAll, 0, 0, "||<Host> A ocurrido un error..." & FONTTYPE_TALK & ENDC)
                Call SendData(ToAll, 0, 0, "||<Host> Cerramos, en 5 min abrimos..." & FONTTYPE_TALK & ENDC)
                Call GuardarUsuarios
                Call SaveGuildsDB
                End
                Exit Sub
            End If
        Next h
        Call SaveGuildsDB
        End
        ' [/GS]
    End If
End If

End Sub
