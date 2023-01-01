Attribute VB_Name = "GS_Party"
' Sistema de partys, version GS
' 16/10/2004
' Para 5 participantes, pero es ampliable

' El en usuario se utilizaran estas 3 variables
'    PartyInvito As Integer ' Ultimo invitado
'    Partys(1 To 5) As Integer
'    LiderParty As Integer

' Idea general:
' Un pj sin estar en una party, puede invitar a otro a formar una party
' El primero que lo haga, es el lider de la party
' De hay solo el lider, podra agregar nuevos integrantes a la party
' Si un integrante decide salir, es sakado sin problemas
' Si el lider decide salir, deshace toda la party y todos son bien informados


' AgregarParty -> EstaEnParty, HayLugarEnParty, DecirATodos
' BorrarParty -> EstaEnParty, EsLiderParty
' DecirATodos -> EstaEnParty, EsLiderParty
' EsLiderParty
' EstaEnParty
' HayLugarEnParty
' InvitarParty -> EsLiderParty, HayLugarEnParty
' PartyExp -> EsLiderParty

Dim i As Integer

Public Function EstaEnParty(ByVal Userindex As Integer) As Boolean
    EstaEnParty = False
    If UserList(Userindex).ConnID <> -1 Then
        If UserList(Userindex).flags.LiderParty <> 0 Then
            ' Si tiene un lider, es un lacayo de la party
            EstaEnParty = True
        Else
            ' Si no tiene lider, miramos si es el Lider el
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) <> 0 Then
                    ' Hay integrantes, entonces es lider de party
                    EstaEnParty = True
                    Exit Function
                End If
            Next
        End If
    Else
        EstaEnParty = False
    End If
End Function

Public Function HayLugarEnParty(ByVal Userindex As Integer) As Boolean
    HayLugarEnParty = False
    If UserList(Userindex).ConnID <> -1 Then
        If UserList(Userindex).flags.LiderParty <> 0 Then
            ' Si tiene un lider, es un lacayo de la party
            HayLugarEnParty = False
        Else
            ' Si no tiene lider, miramos si es el Lider el
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) = 0 Then
                    HayLugarEnParty = True
                    Exit Function
                End If
            Next
        End If
    Else
        HayLugarEnParty = False
    End If
End Function

Public Function InvitarParty(ByVal Userindex As Integer, ByVal tIndex As Integer)
' Userindex = YO
' tIndex = EL

If UserList(tIndex).flags.UserLogged = False Then
    Call SendData(ToIndex, Userindex, 0, "||Debes que clickear sobre un usuario." & FONTTYPE_INFO)
    Exit Function
End If

' Si quien queremos invitar esta en party, no podemos invitarlo
If EstaEnParty(tIndex) Then
    Call SendData(ToIndex, Userindex, 0, "||" & UserList(tIndex).Name & " ya se encuentra en party." & FONTTYPE_INFO)
    Exit Function
End If

If EstaEnParty(Userindex) = False Then ' Si no estoy en party
    ' Si no tengo party, voy a convertirme en lider
    If HayLugarEnParty(Userindex) Then
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te esta invitado a su party. Si deseas aceptar, haz click sobre el y escribe /PARTY." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||La invitacion ha sido enviada." & FONTTYPE_INFO)
        UserList(Userindex).flags.InvitaParty = tIndex
        ' Guardo a quien invite yo!
    Else
        Call SendData(ToIndex, Userindex, 0, "||No hay lugar para invitar mas usuarios." & FONTTYPE_INFO)
    End If
ElseIf EsLiderParty(Userindex) = True Then
    ' Si soy lider
    If HayLugarEnParty(Userindex) Then
        For i = 1 To 5
            If UserList(Userindex).flags.Partys(i) = tIndex Then
                Call SendData(ToIndex, Userindex, 0, "||" & UserList(tIndex).Name & " ya se encuentra en nuestra party." & FONTTYPE_INFO)
                Exit Function
            End If
        Next
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te esta invitado a su party. Si deseas aceptar, haz click sobre el y escribe /PARTY." & FONTTYPE_INFO)
        Call SendData(ToIndex, Userindex, 0, "||La invitacion ha sido enviada." & FONTTYPE_INFO)
        UserList(Userindex).flags.InvitaParty = tIndex
        ' Guardo a quien invite yo!
    Else
        Call SendData(ToIndex, Userindex, 0, "||No hay lugar para invitar mas usuarios." & FONTTYPE_INFO)
    End If
End If

End Function


Public Function AgregarParty(ByVal Userindex As Integer, ByVal tIndex As Integer) As Boolean
    AgregarParty = False
    ' Userindex = yo
    ' tIndex = el
    If EstaEnParty(tIndex) = False Then ' Si el no esta en party
        ' Si el nuevo, no es lider de party y no esta en otra party
        If HayLugarEnParty(Userindex) = False Then ' Si no tengo lugar
            Call SendData(ToIndex, tIndex, 0, "||La party se encuentra completa." & FONTTYPE_INFO)
            Call SendData(ToIndex, Userindex, 0, "||La party se encuentra completa." & FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(Userindex).flags.LiderParty = 0 Then ' Si soy el lider claro
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) = 0 Then ' Busco un lugar limpio
                    If UserList(Userindex).flags.InvitaParty = tIndex Then
                        UserList(Userindex).flags.Partys(i) = tIndex
                        Call SendData(ToIndex, tIndex, 0, "||Ahora perteneces a la party de " & UserList(Userindex).Name & "." & FONTTYPE_INFO)
                        Call SendData(ToIndex, Userindex, 0, "||" & UserList(tIndex).Name & " se a unido a tu party." & FONTTYPE_INFO)
                        UserList(tIndex).flags.LiderParty = Userindex ' yo soy tu lider
                        UserList(Userindex).flags.InvitaParty = 0 ' no no invite a nadie
                        AgregarParty = True
                        Call DecirATodos(tIndex, UserList(tIndex).Name & " se ha unido a nuestra party.")
                    Else
                        Call SendData(ToIndex, tIndex, 0, "||No has sido invitado a esta party." & FONTTYPE_INFO)
                    End If
                    Exit Function
                End If
            Next
        End If
    Else
        Call SendData(ToIndex, tIndex, 0, "||No puedes unirte ahora que estas en una party." & FONTTYPE_INFO)
    End If
End Function

Public Function DecirATodos(ByVal Userindex As Integer, ByVal Mensaje As String)
On Error GoTo Fallo
Dim i As Integer
' userindex yo
If EstaEnParty(Userindex) Then
    If EsLiderParty(Userindex) Then
        Call SendData(ToIndex, Userindex, 0, "||<" & UserList(Userindex).Name & " (Miembro de la Party)> " & Mensaje & "~255~0~255~0~1")
        For i = 1 To 5
            If UserList(Userindex).flags.Partys(i) <> 0 Then
                Call SendData(ToIndex, UserList(Userindex).flags.Partys(i), 0, "||<" & UserList(Userindex).Name & " (Miembro de la Party)> " & Mensaje & "~200~0~255~0~1")
            End If
        Next
    Else ' Integrante
        Call SendData(ToIndex, UserList(Userindex).flags.LiderParty, 0, "||<" & UserList(Userindex).Name & " (Miembro de la Party)> " & (Mensaje) & "~200~0~255~0~1")
        For i = 1 To 5
            If UserList(UserList(Userindex).flags.LiderParty).flags.Partys(i) <> 0 Then
                Call SendData(ToIndex, UserList(UserList(Userindex).flags.LiderParty).flags.Partys(i), 0, "||<" & UserList(Userindex).Name & " (Miembro de la Party)> " & (Mensaje) & "~200~0~255~0~1")
            End If
        Next
    End If
End If
Exit Function
Fallo:
MsgBox Err.Description
End Function

Public Function EsLiderParty(ByVal Userindex As Integer) As Boolean
    EsLiderParty = False
    If UserList(Userindex).flags.LiderParty = 0 Then
        For i = 1 To 5
            If UserList(Userindex).flags.Partys(i) <> 0 Then
                EsLiderParty = True
                Exit Function
            End If
        Next
    End If
End Function

' /dejarparty
Public Function BorrarParty(ByVal Userindex As Integer) As Boolean
    If EstaEnParty(Userindex) Then
        BorrarParty = False
        If EsLiderParty(Userindex) Then
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) <> 0 Then
                    Call SendData(ToIndex, UserList(Userindex).flags.Partys(i), 0, "||" & UserList(Userindex).Name & " ha deshecho la party." & FONTTYPE_INFO)
                    UserList(UserList(Userindex).flags.Partys(i)).flags.LiderParty = 0
                    UserList(Userindex).flags.Partys(i) = 0
                    BorrarParty = True
                End If
            Next
            Call SendData(ToIndex, Userindex, 0, "||Has deshecho tu party." & FONTTYPE_INFO)
            Exit Function
        Else
            If UserList(Userindex).flags.LiderParty <> 0 Then
                For i = 1 To 5
                    If UserList(UserList(Userindex).flags.LiderParty).flags.Partys(i) <> 0 Then
                        Call SendData(ToIndex, UserList(UserList(Userindex).flags.LiderParty).flags.Partys(i), 0, "||" & UserList(Userindex).Name & " ha dejado la party." & FONTTYPE_INFO)
                        UserList(UserList(Userindex).flags.LiderParty).flags.Partys(i) = 0
                        BorrarParty = True
                    End If
                Next
                UserList(Userindex).flags.LiderParty = 0
            End If
        End If
    Else
        Call SendData(ToIndex, Userindex, 0, "||No eres integrante de ninguna party." & FONTTYPE_INFO)
        UserList(Userindex).flags.LiderParty = 0
        UserList(Userindex).flags.InvitaParty = 0
        Exit Function
    End If
End Function

Public Function PartyExp(ByVal Userindex As Integer, ByVal Experiencia As Long)
    Dim CuantosSon As Integer
    ' Userindex = lider
    If Experiencia <= 0 Then Exit Function
    If EstaEnParty(Userindex) Then
        If EsLiderParty(Userindex) Then
            CuantosSon = 1
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) <> 0 Then
                    CuantosSon = CuantosSon + 1
                End If
            Next
            ' Los conto ahora les doy la exp
            Dim EXPcadaUno As Long
            EXPcadaUno = CLng(Experiencia / CuantosSon)
            For i = 1 To 5
                If UserList(Userindex).flags.Partys(i) <> 0 Then
                    Call AddtoVar(UserList(UserList(Userindex).flags.Partys(i)).Stats.exp, EXPcadaUno, MaxExp)
                    Call SendData(ToIndex, UserList(Userindex).flags.Partys(i), 0, "||Has ganado " & EXPcadaUno & " puntos de experiencia." & FONTTYPE_FIGHT_MASCOTA)
                    Call CheckUserLevel(UserList(Userindex).flags.Partys(i))
                End If
            Next
            Call AddtoVar(UserList(Userindex).Stats.exp, EXPcadaUno, MaxExp)
            Call SendData(ToIndex, Userindex, 0, "||Has ganado " & EXPcadaUno & " puntos de experiencia." & FONTTYPE_FIGHT_MASCOTA)
            Call CheckUserLevel(Userindex)
        ElseIf EsLiderParty(Userindex) = False Then ' es un integrante?
            If UserList(Userindex).flags.LiderParty <> 0 Then
                If EsLiderParty(UserList(Userindex).flags.LiderParty) Then
                    Call PartyExp(UserList(Userindex).flags.LiderParty, Experiencia)
                Else
                    UserList(Userindex).flags.LiderParty = 0
                    UserList(Userindex).flags.InvitaParty = 0
                End If
            End If
        End If
    End If
End Function

