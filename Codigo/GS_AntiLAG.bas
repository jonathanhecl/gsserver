Attribute VB_Name = "GS_AntiLAG"
' ############
' NO MORE SHIT
' ############
'
' (r) NMS Optimized
'
' NPS es un nuevo sistema, creado con el unico fin de
' reducir notablemente la comparacion en el microprocesador,
' produciendo una respuesta mas rapida y reduciendo el lag.
' Idea original, EL_OSO
' Programada y Mejorada por ^[GS]^
' Nombre del Sistema, inventado por ^[GS]^

' [GS] Clases Numericas
Public Const CLASS_MAGO = 1
Public Const CLASS_CLERIGO = 2
Public Const CLASS_GUERRERO = 3
Public Const CLASS_ASESINO = 4
Public Const CLASS_LADRON = 5
Public Const CLASS_BARDO = 6
Public Const CLASS_DRUIDA = 7
Public Const CLASS_BANDIDO = 8
Public Const CLASS_PALADIN = 9
Public Const CLASS_CAZADOR = 10
Public Const CLASS_PESCADOR = 11
Public Const CLASS_HERRERO = 12
Public Const CLASS_LEÑADOR = 13
Public Const CLASS_MINERO = 14
Public Const CLASS_CARPINTERO = 15
Public Const CLASS_SASTRE = 16
Public Const CLASS_PIRATA = 17
' [/GS]

' [GS] Razas Numericas
Public Const RAZA_HUMANO = 1
Public Const RAZA_ELFO = 2
Public Const RAZA_ELFO_OSCURO = 3
Public Const RAZA_GNOMO = 4
Public Const RAZA_ENANO = 5
' [/GS]

' [GS] Generos Numericos
Public Const MUJER = 1
Public Const HOMBRE = 2
' [/GS]

' Convierte Genero en Numero
Public Function Gen2Num(ByVal genero As String) As Byte
    On Error Resume Next
    Select Case UCase$(genero)
        Case "MUJER"
            Gen2Num = 1
        Case Else
            Gen2Num = 2
    End Select
End Function

' Convierte de Numero a Genero
Public Function Num2Gen(ByVal genero As Byte) As String
    On Error Resume Next
    Select Case genero
        Case 1
            Num2Gen = "MUJER"
        Case Else
            Num2Gen = "HOMBRE"
    End Select
End Function

' Convierte de Raza a Numero
Public Function Raza2Num(ByVal raza As String) As Byte
    On Error Resume Next
    Raza2Num = 0
    For i = 1 To NUMRAZAS
        If UCase$(ListaRazas(i)) = UCase$(raza) Then
            Raza2Num = i
            Exit Function
        End If
    Next
End Function

' Convierte de Numero a Raza
Public Function Num2Raza(ByVal raza As Byte) As String
    On Error Resume Next
    Num2Raza = ""
    Num2Raza = ListaRazas(raza)
End Function



' Convierte de Clase a Numero
Public Function Clase2Num(ByVal clase As String) As Byte
    On Error Resume Next
    Clase2Num = 0
    For i = 1 To NUMCLASES
        If UCase$(ListaClases(i)) = UCase$(clase) Then
            Clase2Num = i
            Exit Function
        End If
    Next
End Function

' Convierte de Numero a Clase
Public Function Num2Clase(ByVal clase As Byte) As String
    On Error Resume Next
    Num2Clase = ""
    Num2Clase = ListaClases(clase)
End Function



