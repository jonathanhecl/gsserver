Attribute VB_Name = "Mod_General"
'Argentum Online 0.11.2
'
'Copyright (C) 2002 M?rquez Pablo Ignacio
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
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez


Option Explicit

Public bO As Integer
Public bK As Long
Public bRK As Long


Public bKD As Long


Public iplst As String
Public banners As String


Public bInvMod     As Boolean  'El inventario se modific??

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameLimiter As Long

Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Function DirGraficos() As String
DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function
Public Function SD(ByVal N As Integer) As Integer
'Suma digitos
Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal N As Integer) As Integer
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal N As Integer) As Integer

If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If

End Function

Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Sub PlayWaveAPI(File As String)

On Error Resume Next
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Sub CargarAnimArmas()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0

Next loopc

End Sub

Sub CargarVersiones()
On Error GoTo errorH:

Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))

Exit Sub

errorH:
MsgBox ("Error cargando versiones")
End Sub


Sub CargarColores()

Dim archivoC As String
archivoC = App.Path & "\init\colores.dat"

    If Not FileExist(archivoC, vbNormal) Then
        Call MsgBox("ERROR: no se ha podido cargar los colores. falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Integer
    
    For I = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).r = Val(GetVar(archivoC, Str(I), "R"))
        ColoresPJ(I).G = Val(GetVar(archivoC, Str(I), "G"))
        ColoresPJ(I).B = Val(GetVar(archivoC, Str(I), "B"))
    Next I
        
    ColoresPJ(50).r = Val(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = Val(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = Val(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = Val(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = Val(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = Val(GetVar(archivoC, "CI", "B"))
    

End Sub

Sub InitMI()

End Sub

Sub CargarAnimEscudos()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
Next loopc

End Sub

Sub Addtostatus(RichTextBox As RichTextBox, Text As String, RED As Byte, GREEN As Byte, BLUE As Byte, Bold As Byte, Italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

frmCargando.status.SelStart = Len(RichTextBox.Text)
frmCargando.status.SelLength = 0
frmCargando.status.SelColor = RGB(RED, GREEN, BLUE)

If Bold Then
    frmCargando.status.SelBold = True
Else
    frmCargando.status.SelBold = False
End If

If Italic Then
    frmCargando.status.SelItalic = True
Else
    frmCargando.status.SelItalic = False
End If

frmCargando.status.SelText = Chr(13) & Chr(10) & Text

End Sub

Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional RED As Integer = -1, Optional GREEN As Integer, Optional BLUE As Integer, Optional Bold As Boolean, Optional Italic As Boolean, Optional bCrLf As Boolean)
Dim I As Integer
    
    With RichTextBox
        
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = IIf(Bold, True, False)
        .SelItalic = IIf(Italic, True, False)

        If Not RED = -1 Then .SelColor = RGB(RED, GREEN, BLUE)

        .SelText = IIf(bCrLf, Text, Text & vbCrLf)

        RichTextBox.Refresh
    End With
End Sub


Private Function Hex2Dec(ByVal h As String) As Long
Dim I As Long, N As Long, V As Long, C As Long

N = 0
For I = Len(h) To 1 Step -1
    C = Asc(UCase(Mid(h, I, 1)))
    If C >= Asc("A") And C <= Asc("F") Then
        V = C - Asc("A") + 10
    ElseIf C >= Asc("0") And C <= Asc("9") Then
        V = C - Asc("0")
    Else
        V = 0
    End If
    N = N + (16 ^ (Len(h) - I)) * V
Next I

Hex2Dec = N

End Function



'Sub AddtoRichTextBox(RichTextBox As RichTextBox, txt As String, Optional RED As Integer = -1, Optional GREEN As Integer, Optional BLUE As Integer, Optional Bold As Boolean, Optional Italic As Boolean, Optional bCrLf As Boolean)
'Dim i As Long
'Dim N As Long
'Dim Tag As String
'Dim t() As String
'Dim Dale As Boolean
'Dim ColorStack As New Collection
'
'With RichTextBox
'
'If (Len(.Text)) > 2000 Then .Text = ""
'.SelStart = Len(.Text)
'.SelLength = 0
'
''If Not IsMissing(Bold) Then .SelBold = IIf(Bold, True, False)
''If Not IsMissing(Italic) Then .SelItalic = IIf(Italic, True, False)
''If Not IsMissing(RED) And Not IsMissing(GREEN) And Not IsMissing(BLUE) Then .SelColor = RGB(RED, GREEN, BLUE)
'.SelBold = IIf(Bold, True, False)
'.SelItalic = IIf(Italic, True, False)
'If Not RED = -1 Then .SelColor = RGB(RED, GREEN, BLUE)
'
'If InStr(1, txt, "<") > 0 Then
'    i = 1
'    Dale = True
'
'    Do While Dale
'        N = InStr(i, txt, "<")
'        If N > 0 Then
'            .SelText = Mid(txt, i, N - i)
'
'            i = N + 1
'            N = InStr(i, txt, ">")
'            If N > 0 Then
'                Tag = Mid(txt, i, N - i)
'                i = N + 1
'                t = Split(Tag, " ")
'
'                If Len(Tag) > 0 Then
'                    Select Case UCase(t(0))
'                    Case "B"
'                        .SelBold = True
'                    Case "/B"
'                        .SelBold = False
'                    Case "K"
'                        .SelItalic = True
'                    Case "/K"
'                        .SelItalic = False
'                    Case "U"
'                        .SelUnderline = True
'                    Case "/U"
'                        .SelUnderline = False
'                    Case "C"
'                        If UBound(t) > 0 Then
'                            ColorStack.Add .SelColor
'                            .SelColor = IIf(Left(t(1), 1) = "#", Hex2Dec(t(1)), Val(t(1)))
'                        End If
'                    Case "/C"
'                        If ColorStack.Count > 0 Then
'                            .SelColor = ColorStack.Item(ColorStack.Count)
'                            ColorStack.Remove ColorStack.Count
'                        End If
'                    End Select
'                End If
'            Else
'                Dale = False
'            End If
'        Else
'            .SelText = Mid(txt, i)
'            Dale = False
'        End If
'    Loop
'    If Not bCrLf Then .SelText = vbCrLf
'Else
'    .SelText = IIf(bCrLf, txt, txt & vbCrLf)
'End If
'
'.Refresh
'
'End With
'
'End Sub


Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0


TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub

Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************

Dim loopc As Integer

For loopc = 1 To LastChar
    If charlist(loopc).Active = 1 Then
        MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
    End If
Next loopc

End Sub

Sub SaveGameini()
'Grabamos los datos del usuario en el Game.ini

    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort

Call EscribirGameIni(Config_Inicio)

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim I As Integer

cad = LCase$(cad)

For I = 1 To Len(cad)
    car = Asc(Mid$(cad, I, 1))
    
    If ((car < 97 Or car > 122) Or car = Asc("?")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next I

AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean
'Validamos los datos del user
Dim loopc As Integer
Dim CharAscii As Integer

'If IPdelServidor = frmMain.Socket1.LocalAddress Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = "localhost" Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = frmMain.Socket1.LocalName Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = "" Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If PuertoDelServidor = "" Then
'    MsgBox ("Puerto invalido.")
'    Exit Function
'End If

If checkemail Then
 If UserEmail = "" Then
    MsgBox ("Direccion de email invalida")
    Exit Function
 End If
End If

If UserPassword = "" Then
    MsgBox ("Ingrese un password.")
    Exit Function
End If

For loopc = 1 To Len(UserPassword)
    CharAscii = Asc(Mid$(UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Password invalido.")
        Exit Function
    End If
Next loopc

If UserName = "" Then
    MsgBox ("Nombre invalido.")
    Exit Function
End If

If Len(UserName) > 30 Then
    MsgBox ("El nombre debe tener menos de 30 letras.")
    Exit Function
End If

For loopc = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Nombre invalido.")
        Exit Function
    End If
    
Next loopc


CheckUserData = True

End Function
Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

'Set Connected
Connected = True

Call SaveGameini

'Unload the connect form
Unload frmConnect


frmMain.Label8.Caption = UserName
'Load main form
frmMain.Visible = True



End Sub
Sub CargarTip()

Dim N As Integer
N = RandomNumber(1, UBound(Tips))
If N > UBound(Tips) Then N = UBound(Tips)
frmtip.tip.Caption = Tips(N)

End Sub

Sub MoveNorth()
If Cartel Then Cartel = False

If LegalPos(UserPos.X, UserPos.Y - 1) Then
    Call SendData("M" & NORTH)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, NORTH)
        Call MoveScreen(NORTH)
        DoFogataFx
    End If
Else
    If charlist(UserCharIndex).Heading <> NORTH Then
            Call SendData("CHEA" & NORTH)
    End If
End If
End Sub

Sub MoveEast()
If Cartel Then Cartel = False
If LegalPos(UserPos.X + 1, UserPos.Y) Then
    Call SendData("M" & EAST)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, EAST)
        Call MoveScreen(EAST)
        Call DoFogataFx
    End If
Else
    If charlist(UserCharIndex).Heading <> EAST Then
            Call SendData("CHEA" & EAST)
    End If
End If
End Sub

Sub MoveSouth()
If Cartel Then Cartel = False

If LegalPos(UserPos.X, UserPos.Y + 1) Then
    Call SendData("M" & SOUTH)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        MoveCharbyHead UserCharIndex, SOUTH
        MoveScreen SOUTH
        DoFogataFx
    End If
Else
    If charlist(UserCharIndex).Heading <> SOUTH Then
            Call SendData("CHEA" & SOUTH)
    End If
End If
End Sub

Sub MoveWest()
If Cartel Then Cartel = False
If LegalPos(UserPos.X - 1, UserPos.Y) Then
    Call SendData("M" & WEST)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            MoveCharbyHead UserCharIndex, WEST
            MoveScreen WEST
            DoFogataFx
    End If
Else
    If charlist(UserCharIndex).Heading <> WEST Then
            Call SendData("CHEA" & WEST)
    End If
End If
End Sub

Sub RandomMove()

Dim j As Integer

j = RandomNumber(1, 4)

Select Case j
    Case 1
        Call MoveEast
    Case 2
        Call MoveNorth
    Case 3
        Call MoveWest
    Case 4
        Call MoveSouth
End Select

End Sub

Sub CheckKeys()
On Error Resume Next

'*****************************************************************
'Checks keys and respond
'*****************************************************************
Static KeyTimer As Integer

'Makes sure keys aren't being pressed to fast
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If



'Don't allow any these keys during movement..
If UserMoving = 0 Then
    If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveNorth
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveEast
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveSouth
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveWest
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
    Else
        Dim kp As Boolean
        kp = (GetKeyState(vbKeyUp) < 0) Or _
        GetKeyState(vbKeyRight) < 0 Or _
        GetKeyState(vbKeyDown) < 0 Or _
        GetKeyState(vbKeyLeft) < 0
        If kp Then Call RandomMove
        If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
        frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
    End If
End If

End Sub




Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1

    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
Exit Sub
Stop
    '[CODE 001]:MatuX'
        ' Frame checker para el cheat ese
        Select Case FramesPerSecCounter
            Case 18 To 19
                lFrameModLimiter = 60
            Case 17
                lFrameModLimiter = 60
            Case 16
                lFrameModLimiter = 120
            Case 15
                lFrameModLimiter = 240
            Case 14
                lFrameModLimiter = 480
            Case 15
                lFrameModLimiter = 960
            Case 14
                lFrameModLimiter = 1920
            Case 13
                lFrameModLimiter = 3840
            Case 12
            Case 11
            Case 10
            Case 9
            Case 8
            Case 7
            Case 6
            Case 5
            Case 4
            Case 3
            Case 2
            Case 1
                lFrameModLimiter = 60 * 256
            Case 0
            
        End Select
    '[END]'

    Call DoFogataFx
End If

End Sub

Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************

Dim loopc As Integer

loopc = 1
Do While charlist(loopc).Active And loopc < UBound(charlist)
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Public Function DirMapas() As String
DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Sub SwitchMap(Map As Integer)

Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer
      

Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
Seek #1, 1
        
'map Header
Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
        
'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        '.dat file
        Get #1, , MapData(X, Y).Blocked
        For loopc = 1 To 4
            Get #1, , MapData(X, Y).Graphic(loopc).GrhIndex
            
            'Set up GRH
            If MapData(X, Y).Graphic(loopc).GrhIndex > 0 Then
                InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).GrhIndex
            End If
            
        Next loopc
        
        
        Get #1, , MapData(X, Y).Trigger
        
        Get #1, , tempint
        
        'Erase NPCs
        If MapData(X, Y).CharIndex > 0 Then
            Call EraseChar(MapData(X, Y).CharIndex)
        End If
        
        'Erase OBJs
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

Close #1

MapInfo.Name = ""
MapInfo.Music = ""

CurMap = Map

End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim I As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For I = 1 To Len(Text)
    CurChar = Mid(Text, I, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = I
    End If
Next I
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Sub WriteClientVer()

Dim hFile As Integer
    
hFile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hFile
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)

Put #hFile, , CInt(App.Major)
Put #hFile, , CInt(App.Minor)
Put #hFile, , CInt(App.Revision)

Close #hFile

End Sub


Public Function IsIp(ByVal Ip As String) As Boolean

Dim I As Integer
For I = 1 To UBound(ServersLst)
    If ServersLst(I).Ip = Ip Then
        IsIp = True
        Exit Function
    End If
Next I

End Function

Public Sub CargarServidores()
On Error GoTo errorH
Dim f As String
Dim C As Integer
Dim I As Integer

f = App.Path & "\init\sinfo.dat"
C = Val(GetVar(f, "INIT", "Cant"))

ReDim ServersLst(1 To C) As tServerInfo
For I = 1 To C
    ServersLst(I).desc = GetVar(f, "S" & I, "Desc")
    ServersLst(I).Ip = Trim(GetVar(f, "S" & I, "Ip"))
    ServersLst(I).PassRecPort = Val(GetVar(f, "S" & I, "P2"))
    ServersLst(I).Puerto = Val(GetVar(f, "S" & I, "PJ"))
Next I
CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
End Sub

Public Sub InitServersList(ByVal Lst As String)

On Error Resume Next

Dim NumServers As Integer
Dim I As Integer, Cont As Integer
I = 1

Do While (ReadField(I, RawServersList, Asc(";")) <> "")
    I = I + 1
    Cont = Cont + 1
Loop

ReDim ServersLst(1 To Cont) As tServerInfo

For I = 1 To Cont
    Dim cur$
    cur$ = ReadField(I, RawServersList, Asc(";"))
    ServersLst(I).Ip = ReadField(1, cur$, Asc(":"))
    ServersLst(I).Puerto = ReadField(2, cur$, Asc(":"))
    ServersLst(I).desc = ReadField(4, cur$, Asc(":"))
    ServersLst(I).PassRecPort = ReadField(3, cur$, Asc(":"))
Next I

CurServer = 1



End Sub

Public Function CurServerPasRecPort() As Integer

If CurServer <> 0 Then
'    CurServerPasRecPort = ServersLst(CurServer).PassRecPort
    CurServerPasRecPort = 7667
    
Else
    CurServerPasRecPort = CInt(frmConnect.PortTxt)
End If

End Function


Public Function CurServerIp() As String

If CurServer <> 0 Then
    CurServerIp = ServersLst(CurServer).Ip
Else
    CurServerIp = frmConnect.IPTxt
End If

End Function

Public Function CurServerPort() As Integer

If CurServer <> 0 Then
    CurServerPort = ServersLst(CurServer).Puerto
Else
    CurServerPort = CInt(frmConnect.PortTxt)
End If

End Function


Sub Main()
On Error Resume Next

ChDir App.Path

Call WriteClientVer

Call LeerLineaComandos

Dim cRes As Integer
cRes = MsgBox("? Jugar a tama?o pantalla 800x600 ?", vbYesNoCancel + vbQuestion)
If cRes = vbYes Then
    NoRes = False
Else
    NoRes = True
End If

If App.PrevInstance Then
    Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Integer

ChDrive App.Path
ChDir App.Path

'Obtengo mi MD5 hash
'Obtener el HushMD5
Dim fMD5HushYo As String * 32
fMD5HushYo = MD5File(App.Path & "\" & App.EXEName & ".exe")
'fMD5HushYo = MD5File(App.Path & "\ARGENTUM.exe")
'fMD5HushYo = MD5File(App.Path & "\" & "argentumdinamicotest.exe")

MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 53)


'Cargamos el archivo de configuracion inicial
If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
    Config_Inicio = LeerGameIni()
End If


If FileExist(App.Path & "\init\ao.dat", vbNormal) Then
    Open App.Path & "\init\ao.dat" For Binary As #53
        Get #53, , RenderMod
    Close #53

    Musica = IIf(RenderMod.bNoMusic = 1, 1, 0)
    Fx = IIf(RenderMod.bNoSound = 1, 1, 0)
    
    'RenderMod.iImageSize = 0
    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End If


tipf = Config_Inicio.tip

frmCargando.Show
frmCargando.Refresh

UserParalizado = False

frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
AddtoRichTextBox frmCargando.status, "Buscando servidores....", 0, 0, 0, 0, 0, 1


'If RawServersList = "" Then
'    frmMain.Inet1.URL = "http://www.argentum-online.com.ar/admin/iplist2.txt"
'End If

#If UsarWrench = 1 Then

frmMain.Socket1.Startup

#Else

#End If

Call CargarServidores
ServersRecibidos = True
'IPdelServidor =
'PuertoDelServidor = 7666

AddtoRichTextBox frmCargando.status, "Encontrado", , , , 1
AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1

ReDim Ciudades(1 To NUMCIUDADES) As String
Ciudades(1) = "Ullathorpe"
Ciudades(2) = "Nix"
Ciudades(3) = "Banderbill"

ReDim CityDesc(1 To NUMCIUDADES) As String
CityDesc(1) = "Ullathorpe est? establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y le?adores. Su ubicaci?n hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares m?s legendarios de este mundo."
CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades m?s importantes de todo el imperio."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"



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
ListaClases(13) = "Le?ador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apu?alar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar ?rboles"
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
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"


'CLSCONTENEDORCHARFILES




frmOldPersonaje.NameTxt.Text = Config_Inicio.Name
frmOldPersonaje.PasswordTxt.Text = ""

AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

IniciarObjetosDirectX

AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 0, 0, 0, 0, 1
AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

Dim loopc As Integer

LastTime = GetTickCount

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

Call InitTileEngine(frmMain.hWnd, 152, 7, 32, 32, 13, 17, 9)
                                  

'Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extras.", 2, 51, 223, 1, 1)
Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra....")


Call CargarAnimsExtra
Call CargarTips
UserMap = 1
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarVersiones
Call CargarColores


AddtoRichTextBox frmCargando.status, "                    ?Bienvenido a Argentum Online!", , , , 1


Unload frmCargando

LoopMidi = True

If Musica = 0 Then
    Call CargarMIDI(DirMidi & MIdi_Inicio & ".mid")
    Play_Midi
End If

frmPres.Picture = LoadPicture(App.Path & "\Graficos\bosquefinal.jpg")
'frmPres.WindowState = vbMaximized
frmPres.Show

Do While Not finpres
    DoEvents
Loop

Unload frmPres

frmConnect.Visible = True

'Loop principal!
'[CODE]:MatuX'
    MainViewRect.Left = MainViewLeft + 32 * RenderMod.iImageSize
    MainViewRect.Top = MainViewTop + 32 * RenderMod.iImageSize
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    MainDestRect.Left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
    MainDestRect.Top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
    MainDestRect.Right = (MainDestRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainDestRect.Bottom = (MainDestRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
'[END]'


Dim mainAntX As Long, mainAntY As Long

PrimeraVez = True
prgRun = True
pausa = False
bInvMod = True
lFrameLimiter = DirectX.TickCount
'[CODE 001]:MatuX'
    lFrameModLimiter = 60
'[END]'
Do While prgRun

    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Pedimos que nos envie la posicion
            Call SendData("RPU")
        End If
    End If

'    Call RefreshAllChars

    '[CODE 001]:MatuX
    '
    '   EngineRun
    If EngineRun Then
        '[DO]:Dibuja el siguiente frame'
        '[CODE 000]:MatuX'
        'If frmMain.WindowState <> 1 And CurMap > 0 And EngineRun Then
        If frmMain.WindowState <> 1 Then
        '[END]'
            'Call ShowNextFrame(frmMain.Top, frmMain.Left)
            '****** Move screen Left, Right, Up and Down if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If
    
            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            'Call DoNightFX
            'Call DoLightFogata(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
            '[CODE 000]:MatuX
                'Call MostrarFlags
                If IScombate Then Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed)
                If Dialogos.CantidadDialogos <> 0 Then Call Dialogos.MostrarTexto
                If Cartel Then Call DibujarCartel
                If bInvMod Then DibujarInv
                
                If mainAntX <> frmMain.Left Or mainAntY <> frmMain.Top Then
                    mainAntX = frmMain.Left
                    mainAntY = frmMain.Top
                    MainViewRect.Left = (frmMain.Left / Screen.TwipsPerPixelX) + MainViewLeft + 32 * RenderMod.iImageSize
                    MainViewRect.Top = (frmMain.Top / Screen.TwipsPerPixelY) + MainViewTop + 32 * RenderMod.iImageSize
                    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
                    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
                End If
                
                Call DrawBackBufferSurface
                
                Call RenderSounds
                
                '[DO]:Inventario'
                'Call DibujarInv(frmMain.picInv.hWnd, 0)
                'If bInvMod Then DibujarInv  'lo mov? arriba para
                '                             que est? mas ordenadito
                '[END]'
    
            '[END]'
            
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    
    '[CODE 000]:MatuX'
    'If ControlVelocidad(LastTime) Then
    If (GetTickCount - LastTime > 20) Then
        If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
            CheckKeys
            LastTime = GetTickCount
        End If
    End If
    
    If Musica = 0 Then
        If Not SegState Is Nothing Then
            If Not Perf.IsPlaying(Seg, SegState) Then Play_Midi
        End If
    End If
         'Musica = 0
    'End If
    '[END]'
    
    '[CODE 001]:MatuX
    ' Frame Limiter
        'FramesPerSec = FramesPerSec + 1
        If DirectX.TickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            If FPSFLAG Then frmMain.Caption = FramesPerSec
            FramesPerSecCounter = 0
            lFrameTimer = DirectX.TickCount
        End If
        
        'While DirectX.TickCount - lFrameLimiter < lFrameModLimiter: Wend
        
        '[Alejo]
            While DirectX.TickCount - lFrameLimiter < 55 '< 55
                Sleep 5
            Wend
        '[/Alejo]
        

        lFrameLimiter = DirectX.TickCount
    
    '[END]'
    
    'Sistema de timers renovado:
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
        timers(loopc) = timers(loopc) + (esttick - ulttick)
        'timer de trabajo
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
        'timer de attaque (77)
        If timers(2) >= tAt Then
            timers(2) = 0
            UserCanAttack = 1
            UserPuedeRefrescar = True
        End If
    Next loopc
    ulttick = GetTickCount
    
    
       DoEvents
Loop

EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
LiberarObjetosDX


If bNoResChange = False Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
    
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
           .dmPelsHeight = oldResHeight
        End With
lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
End If


Call UnloadAllForms

Config_Inicio.tip = tipf
Call EscribirGameIni(Config_Inicio)

End

ManejadorErrores:
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
    
End Sub



Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function


'[CODE 002]:MatuX
'
'  Funci?n para chequear el email
'
    Public Function CheckMailString(ByRef sString As String) As Boolean
        On Error GoTo errHnd:
        Dim lPos  As Long, lX    As Long
        Dim iAsc  As Integer
    
        '1er test: Busca un simbolo @
        lPos = InStr(sString, "@")
        If (lPos <> 0) Then
            '2do test: Busca un simbolo . despu?s de @ + 1
            If Not (IIf((InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1)), True, False)) Then _
                Exit Function
    
            '3er test: Val?da el ultimo caracter
            If Not (CMSValidateChar_(Asc(Right(sString, 1)))) Then _
                Exit Function
    
            '4to test: Recorre todos los caracteres y los val?da
            For lX = 0 To Len(sString) - 1 'el ultimo no porque ya lo probamos
                If Not (lX = (lPos - 1)) Then
                    iAsc = Asc(Mid(sString, (lX + 1), 1))
                    If Not (iAsc = 46 And lX > (lPos - 1)) Then _
                        If Not CMSValidateChar_(iAsc) Then _
                            Exit Function
                End If
            Next lX
    
            'Finale
            CheckMailString = True
        End If
    
errHnd:
        'Error Handle
    End Function
    
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean
CMSValidateChar_ = IIf( _
                    (iAsc >= 48 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95) Or (iAsc = 45), True, False)
End Function


Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
   MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
   MapData(X, Y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function



    Public Sub ShowSendTxt()
        If Not frmCantidad.Visible Then
            frmMain.SendTxt.Visible = True
            frmMain.SendTxt.SetFocus
        End If
    End Sub
    Public Sub ShowSendCMSGTxt()
        If Not frmCantidad.Visible Then
            frmMain.SendCMSTXT.Visible = True
            frmMain.SendCMSTXT.SetFocus
        End If
    End Sub
    


Public Sub LeerLineaComandos()
Dim Tmp As String, T() As String
Dim I As Long

'inicializo los parametros estandar
NoRes = False 'si esta en false, la cambio

Tmp = Command
T = Split(Tmp, " ")

I = LBound(T)
Do While I <= UBound(T)
    Select Case UCase(T(I))
    Case "/NORES" 'no cambiar la resolucion
        NoRes = True
    End Select
    I = I + 1
Loop

End Sub

Public Function Reverse(ByVal s As String) As String
Dim I As Integer
Reverse = vbNullString

    For I = 1 To Len(s)
        Reverse = Mid$(s, I, 1) & Reverse
    Next I
End Function
