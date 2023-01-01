Attribute VB_Name = "mdlCOmercioConUsuario"
'Modulo para comerciar con otro usuario
'Por Alejo (Alejandro Santos)
'
'
'[Alejo]

Option Explicit

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    Cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(Origen As Integer, Destino As Integer)
On Error GoTo errhandler

'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(ToIndex, Origen, 0, "INITCOMUSU")
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(ToIndex, Destino, 0, "INITCOMUSU")
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call SendData(ToIndex, Destino, 0, "||" & UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Exit Sub
errhandler:

End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(AQuien As Integer)
'Dim Object As UserOBJ
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

'Object.Amount = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    'Object.ObjIndex = iORO
    ObjInd = iORO
Else
    'Object.ObjIndex = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

'If Object.ObjIndex > 0 And Object.Amount > 0 Then
'    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
'    & ObjData(Object.ObjIndex).ObjType & "," _
'    & ObjData(Object.ObjIndex).MaxHIT & "," _
'    & ObjData(Object.ObjIndex).MinHIT & "," _
'    & ObjData(Object.ObjIndex).MaxDef & "," _
'    & ObjData(Object.ObjIndex).Valor \ 3)
'End If
If ObjInd > 0 And ObjCant > 0 Then
    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & ObjInd & "," & ObjData(ObjInd).Name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).ObjType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3)
End If

End Sub

Public Sub FinComerciarUsu(Userindex As Integer)
UserList(Userindex).ComUsu.Acepto = False
UserList(Userindex).ComUsu.Cant = 0
UserList(Userindex).ComUsu.DestUsu = 0
UserList(Userindex).ComUsu.Objeto = 0

UserList(Userindex).flags.Comerciando = False

Call SendData(ToIndex, Userindex, 0, "FINCOMUSUOK")
End Sub

Public Sub AceptarComercioUsu(Userindex As Integer)
If UserList(Userindex).ComUsu.DestUsu <= 0 Or _
    UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.DestUsu <> Userindex Then
    Exit Sub
End If

UserList(Userindex).ComUsu.Acepto = True

If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    Call SendData(ToIndex, Userindex, 0, "||El otro usuario aun no ha aceptado tu oferta." & FONTTYPE_TALK)
    Exit Sub
End If

Dim Obj1 As Obj, Obj2 As Obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False
OtroUserIndex = UserList(Userindex).ComUsu.DestUsu

'[Alejo]: Creo haber podido erradicar el bug de
'         no poder comerciar con mas de 32k de oro.
'         Las lineas comentadas en los siguientes
'         2 grandes bloques IF (4 lineas) son las
'         que originaban el problema.

If UserList(Userindex).ComUsu.Objeto = FLAGORO Then
    'Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = iORO
    'If Obj1.Amount > UserList(UserIndex).Stats.GLD Then
    If UserList(Userindex).ComUsu.Cant > UserList(Userindex).Stats.GLD Then
        Call SendData(ToIndex, Userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(Userindex).ComUsu.Cant
    Obj1.ObjIndex = UserList(Userindex).Invent.Object(UserList(Userindex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(Userindex).Invent.Object(UserList(Userindex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, Userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = iORO
    'If Obj2.Amount > UserList(OtroUserIndex).Stats.GLD Then
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(Userindex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

'[CORREGIDO]
'Desde acá corregí el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

If ObjData(Obj2.ObjIndex).NoSePasa = True Then
    Call SendData(ToIndex, Userindex, 0, "||Este objeto no puede ser vendido." & FONTTYPE_INFO)
    GoTo Finale
End If

If ObjData(Obj1.ObjIndex).NoSePasa = True Then
    Call SendData(ToIndex, OtroUserIndex, 0, "||Este objeto no puede ser vendido." & FONTTYPE_INFO)
    GoTo Finale
End If

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserStatsBox(OtroUserIndex)
    'y se la doy al otro
    UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserStatsBox(Userindex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(Userindex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If

'pone el oro directamente en la billetera
If UserList(Userindex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - UserList(Userindex).ComUsu.Cant
    Call SendUserStatsBox(Userindex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(Userindex).ComUsu.Cant
    Call SendUserStatsBox(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, Userindex)
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, Userindex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Finale:

Call FinComerciarUsu(Userindex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]

