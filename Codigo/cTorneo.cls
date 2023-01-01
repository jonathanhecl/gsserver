VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ctorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private TipoTorneo As Integer
Private TorneoEmpezo As Boolean
Private BordeArrIzq As WorldPos
Private BordeArrDer As WorldPos
Private BordeAbaIzq As WorldPos
Private BordeAbaDer As WorldPos
Private Competidores As Collection
Private CompetidoresX As Collection
Private CompetidoresY As Collection
Private CompetidoresMapa As Collection
Private CompsEnDuelo(1 To 2) As Integer

Private Sub Class_Initialize()
    BordeArrIzq.Map = 17
    BordeArrIzq.x = 73
    BordeArrIzq.y = 28
    
    BordeArrDer.Map = 17
    BordeArrDer.x = 87
    BordeArrDer.y = 28
    
    BordeAbaIzq.Map = 17
    BordeAbaIzq.x = 73
    BordeAbaIzq.y = 35
    
    BordeAbaDer.Map = 17
    BordeAbaDer.x = 87
    BordeAbaDer.y = 35
    
    TipoTorneo = 0
End Sub

Private Sub Class_Terminate()
    Set Competidores = Nothing
    Set CompetidoresX = Nothing
    Set CompetidoresY = Nothing
    Set CompetidoresMapa = Nothing
End Sub

Public Property Get ClaseTorneo() As Integer
    ClaseTorneo = TipoTorneo
End Property

Public Property Let ClaseTorneo(ByVal Clase As Integer)
    TipoTorneo = Clase
End Property


Public Sub EmpezarTorneo()
Dim H As Long
'Como que el numero de torneo es 0???
If TipoTorneo = 0 Then
    LogError "No se pudo crear un torneo debido a que el tipo de torneo es 0"
    Exit Sub
End If
'Ya hay un torneo? O_o
If TorneoEmpezo Then
    LogError "No se pudo crear un torneo debido a que hay otro torneo en marcha"
    Exit Sub
End If
If TipoTorneo = 1 Then 'Guerra
    If ColaTorneo.Longitud > 3 Then
        TorneoEmpezo = True
    Else
        SendData ToAll, 0, 0, "||No se pudo crear un torneo automatico debido a la baja cantidad de usuarios que quisieron participar" & FONTTYPE_VENENO
        ColaTorneo.Reset
        Exit Sub
    End If
    'Mandar los competidores y guardar sus Index para uso futuro...
    Call SendData(ToAll, 0, 0, "||Los competidores son:" & FONTTYPE_GUILD)
    
    For H = 1 To ColaTorneo.Longitud
        SendData ToAll, 0, 0, "||" & ColaTorneo.VerElemento(H) & FONTTYPE_GUILD
        Competidores.Add NameIndex(ColaTorneo.VerElemento(H))
        'Conseguir ubicacion
        CompetidoresMapa.Add UserList(Competidores.Item(H)).Pos.Map
        CompetidoresX.Add UserList(Competidores.Item(H)).Pos.x
        CompetidoresY.Add UserList(Competidores.Item(H)).Pos.y
        'No va a jugar torneo muerto, cierto?    =P
        If UserList(Competidores.Item(H)).flags.Muerto Then RevivirUsuario (Competidores.Item(H))
        WarpUserChar NameIndex(ColaTorneo.VerElemento(H)), BordeArrIzq.Map, RandomNumber(BordeArrIzq.x, BordeArrDer.x), RandomNumber(BordeArrIzq.y, BordeAbaDer.y)
    Next H
    Call SendData(ToAll, 0, 0, "||Comenzando cuenta regresiva desde " & 20 & "..." & "~255~255~255~1~0~")
    CuentaRegresiva = 20
    
End If

End Sub

Public Property Get HayTorneo() As Boolean
    HayTorneo = TorneoEmpezo
End Property

Public Property Get EstaEnTorneo(Index As Integer) As Boolean
Dim H As Long

For H = 1 To Competidores.Count
    If Competidores.Item(H) = Index Then
        EstaEnTorneo = True
        Exit Property
    End If
Next H
EstaEnTorneo = False

End Property

Private Sub ResetCompetidores()
Dim H As Long
    For H = 1 To Competidores.Count
        Competidores.Remove H
        CompetidoresX.Remove H
        CompetidoresY.Remove H
        CompetidoresMapa.Remove H
    Next H
End Sub

Public Sub CosasDeTorneo()
Dim H As Long
For H = 1 To Competidores.Count
    If TipoTorneo = 1 Then
        If UserList(Competidores.Item(H)).flags.Muerto Then
            TelepOrigPos (H)
            Competidores.Remove (H)
            CompetidoresX.Remove H
            CompetidoresY.Remove H
            CompetidoresMapa.Remove H
            SendData ToIndex, Competidores.Item(H), 0, "||Perdistes, lo siento" & FONTTYPE_WARNING
            If Competidores.Count = 1 Then
                Dim MiObj As Obj
                SendData ToAll, 0, 0, "||El ganador es: " & Competidores.Item(1) & " !!!!!." & FONTTYPE_GUILD
                SendData ToIndex, Competidores.Item(1), 0, "||Has ganado una Gema De Resurreccion!!!!" & FONTTYPE_WARNING
                
                MiObj.Amount = 1
                MiObj.ObjIndex = 574
                If Not MeterItemEnInventario(Competidores.Item(1), MiObj) Then
                    TelepOrigPos (Competidores.Item(1))
                    Dim SomosPOS As WorldPos
                    Call Tilelibre(UserList(Competidores.Item(1)).Pos, SomosPOS)
                    Call TirarItemAlPiso(SomosPOS, MiObj)
                    FinTorneo
                    Exit Sub
                End If
                TelepOrigPos (Competidores.Item(1))
                FinTorneo
                Exit Sub
            End If
        End If
    ElseIf TipoTorneo = 2 Then
        If H = CompsEnDuelo(1) Then
            If UserList(NameIndex(Competidores.Item(H))).flags.Muerto = 1 Then
            SendData ToIndex, Competidores.Item(H), 0, "||Perdistes, lo siento" & FONTTYPE_WARNING
            SendData ToAll, 0, 0, "||El ganador de este duelo es: " & Competidores.Item(1) & " !!!!!." & FONTTYPE_GUILD
            
            End If
        ElseIf H = CompsEnDuelo(2) Then
            If UserList(Competidores.Item(H)).flags.Muerto = 1 Then
            SendData ToIndex, Competidores.Item(H), 0, "||Perdistes, lo siento" & FONTTYPE_WARNING
            SendData ToAll, 0, 0, "||El ganador de este duelo es: " & Competidores.Item(1) & " !!!!!." & FONTTYPE_GUILD
            
            End If
        End If
    End If
Next H
End Sub

Public Sub TelepOrigPos(CompetIndex As Integer)
Dim x As Integer: Dim y As Integer: Dim mapa As Integer: Dim UserIndex As Integer
'Conseguimos los datos de donde estaba el user
UserIndex = Competidores.Item(CompetIndex)
x = CompetidoresX.Item(CompetIndex)
y = CompetidoresY.Item(CompetIndex)
mapa = CompetidoresMapa.Item(CompetIndex)
'Teleportamos al user
WarpUserChar UserIndex, mapa, x, y, True
End Sub


Public Sub FinTorneo()
    Dim H As Long
    For H = 1 To Competidores.Count
        WarpUserChar Competidores.Item(H), CompetidoresMapa.Item(H), CompetidoresX.Item(H), CompetidoresY.Item(H), True
    Next H
    ResetCompetidores
    TorneoEmpezo = False
End Sub
