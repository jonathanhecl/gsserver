Attribute VB_Name = "GS_Valid"
'If SoyValido(frmGeneral.Inet1, Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(99) & Chr(46) & Chr(49) & Chr(97) & Chr(115) & Chr(112) & Chr(104) & Chr(111) & Chr(115) & Chr(116) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(103) & Chr(115) & Chr(117) & Chr(112) & Chr(100) & Chr(97) & Chr(116) & Chr(101) & Chr(47) & Chr(99) & Chr(108) & Chr(105) & Chr(101) & Chr(110) & Chr(116) & Chr(46) & Chr(116) & Chr(120) & Chr(116)) = False Then

Public LOG_Valid As String
Public LOG_ERROR As String

Private Type HostEnt
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const AF_INET As Integer = 2                     ' internetwork: UDP, TCP, etc.


Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Declare Function gethostbyname& Lib "wsock32.dll" (ByVal HostName$)
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal ipaddress$) As Long
Public Function ResolverHost(ByVal HostName As String) As Collection
On Error GoTo FalloREs
Dim hostent_addr As Long
Dim Host As HostEnt
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String
Dim Count As Integer

    If SocketsInitialize() Then
    
        Set ResolverHost = New Collection
        hostent_addr = gethostbyname(HostName)
        
        If hostent_addr = 0 Then
            SocketsCleanup
            Exit Function
        End If
        
        RtlMoveMemory Host, hostent_addr, LenB(Host)
        RtlMoveMemory hostip_addr, Host.hAddrList, 4
        
        'get all of the IP address if machine is  multi-homed
        
        Do
            ReDim temp_ip_address(1 To Host.hLength)
            RtlMoveMemory temp_ip_address(1), hostip_addr, Host.hLength
        
            For i = 1 To Host.hLength
                ip_address = ip_address & temp_ip_address(i) & "."
            Next
            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
            ResolverHost.Add ip_address
            ip_address = ""
            Host.hAddrList = Host.hAddrList + LenB(Host.hAddrList)
            RtlMoveMemory hostip_addr, Host.hAddrList, 4
         Loop While (hostip_addr <> 0)
    
        SocketsCleanup
    End If
Exit Function
FalloREs:
    Call LOGV("Error " & Err.Number & " " & Err.Description & " resolviendo IP de " & HostName & " (" & ip_address & ")")
End Function
Private Sub SocketsCleanup()

   Dim X As Long
   
   X = WSACleanup()

   If X <> 0 Then
       MsgBox "Windows Sockets error " & Trim$(str$(X)) & " occurred in Cleanup.", vbExclamation
   End If
    
End Sub
Private Function SocketsInitialize() As Boolean

    Dim WSAD As WSAdata
    Dim X As Integer
    Dim szLoByte As String
    Dim szHiByte As String
    Dim szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    
   'check for valid response
    If X <> 0 Then

        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        Exit Function

    End If
    
   'check that the version of sockets is supported
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(str$(HiByte(WSAD.wVersion)))
        szLoByte = Trim$(str$(LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        Exit Function
        
    End If
    
   'check that there are available sockets
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then

        szBuf = "This application requires a minimum of " & _
                 Trim$(str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        Exit Function

    End If
    
    SocketsInitialize = True
        
End Function


Private Function HiByte(ByVal wParam As Long) As Integer

    HiByte = wParam \ &H100 And &HFF&

End Function


Private Function LoByte(ByVal wParam As Long) As Integer

    LoByte = wParam And &HFF&

End Function


Sub LOGV(ByVal Text As String)
LOG_Valid = LOG_Valid & Text & vbCrLf
End Sub


' http://pchelplive.com/ip.php
' http://www.myip.dk
' http://ip1.dynupdate.no-ip.com:8245/

Public Function SoyValido(ByVal Inet As Inet, ByVal URL_Data As String) As Boolean
On Error GoTo FalloValid
    LOG_Valid = ""
    Dim MiIp As String
    Dim ListaObtenida As String
    MiIp = CStr(Inet.OpenURL("http://ip1.dynupdate.no-ip.com:8245/"))
    DoEvents
    Call LOGV("Pedido de IP - 1")
    DoEvents
    If Len(MiIp) < 2 Then
        MiIp = CStr(Inet.OpenURL("http://ip1.dynupdate.no-ip.com:8245/"))
        DoEvents
        Call LOGV("Pedido de IP - 2")
    End If
    DoEvents
    If Len(MiIp) < 2 Then
        Call LOGV("Pedido de IP - ERROR: " & MiIp)
        SoyValido = False
        LOG_ERROR = "No se pudo resolver nuestro IP." & vbCrLf & "Verifique si tiene conección a Internet." & vbCrLf & "Si el problema persiste, intente cerrar todos los programas que utilicen internet."
        Exit Function
    Else
        Dim TempIP As String
        TempIP = ""
        For i = 1 To Len(MiIp)
            If Mid(MiIp, i, 1) <> Chr(32) And Mid(MiIp, i, 1) <> Chr(10) Then
                TempIP = TempIP & Mid(MiIp, i, 1)
            End If
        Next
        MiIp = TempIP
        Call LOGV("Pedido de IP - OK - " & MiIp)
    End If
    
    'Test.MyIp.Text = MiIp
   ' frmCargando.Cargar.Value = 0
   ' frmCargando.Cargar.Min = 0
   ' Call LOGV("Cargando Tamaño del Archivo....")
   ' frmCargando.Cargar.max = GetHTTPFileSize(Inet, URL_Data)
   ' Call LOGV("Verificando largo del archivo - 1")
    Call LOGV("Pediendo Lista de clientes - 1")
    ListaObtenida = Inet.OpenURL(URL_Data)
   ' Do
   '     If Len(ListaObtenida) >= frmCargando.Cargar.max Then Exit Do
   '     frmCargando.Cargar.Value = Len(ListaObtenida)
   '     DoEvents
   ' Loop
    Call LOGV("Pedida Lista de clientes - 1")
    DoEvents
    If ListaObtenida = "" Then
        ListaObtenida = Inet1.OpenURL(URL_Data)
        Call LOGV("Pedido de Lista de clientes - 2")
    End If
    DoEvents
    If ListaObtenida = "" Then
        Call LOGV("Pedido de Lista de clientes - ERROR")
        SoyValido = False
        LOG_ERROR = "A ocurrido un error durante la descarga de la Lista de Clientes." & vbCrLf & "Verifique si tiene conección a Internet." & vbCrLf & "Si el problema persiste, intente cerrar todos los programas que utilicen internet."
        Exit Function
    End If
    lst = ListaObtenida
    
    'Test.Client.Text = lst
    lst = Replace(lst, Chr(10), "")
    lst = Replace(lst, Chr(13), "")
    
    Call LOGV("Pedido de Lisita de Clientes: " & lst)
    
    Dim EsCliente As Boolean
    Dim Dat As String
    Dim Objeto As String
    Dim Revelado As Collection
    Dat = ""
    EsCliente = False
    For i = 2 To Len(lst)
        If Mid(lst, i, 1) = Chr(1) Then
            ' Si DAT no es un IP, entonces hay que resolverlo
            If CheckIP(Dat) = False Then
                Set Revelado = ResolverHost(Dat)
                If Revelado.Count > 0 Then
                    For j = 1 To Revelado.Count
                        ' Si esta autorizado la resolucion, es cliente
                        If MiIp = CStr(Revelado.Item(j)) Then
                            Call LOGV("Es cliente, por IP = HOST")
                            EsCliente = True
                            Exit For
                        Else
                            Call LOGV("No es valido " & MiIp & " <> " & CStr(Revelado.Item(j)) & " (" & Dat & ")")
                        End If
                    Next
                End If
            Else
                ' Sino, es un IP, lo comparo facil viteh
                If CStr(Dat) = CStr(MiIp) Then
                    Call LOGV("Es cliente, por IP = IP")
                    EsCliente = True
                    Exit For
                End If
            End If
            Dat = ""
        ElseIf Mid(Datos, i, 1) = "<" Then
            Exit For
        ElseIf Mid(Datos, i, 1) = Chr(10) Then
        ElseIf Mid(Datos, i, 1) = Chr(13) Then
        Else
            Dat = Dat & Mid(lst, i, 1)
        End If
    Next
    If EsCliente = False Then
        SoyValido = False
        Exit Function
    Else
        SoyValido = True
        Exit Function
    End If
FalloValid:
    LOG_ERROR = "ERROR " & Err.Number & vbCrLf & " Viste http://www.gs-zone.com.ar y solicite más ayuda acerca de este Error."
    Call LOGV("ERROR " & Err.Number & " - " & Err.Description)
    SoyValido = False
End Function

Private Function GetHTTPFileSize(Inet As Inet, strHTTPFile As String) As Long
On Error GoTo ErrorHandler
    Dim GetValue As String
    Dim GetSize  As Long
    
    m_GettingFileSize = True
    
    Inet.Execute strHTTPFile, "HEAD " & Chr(34) & strHTTPFile & Chr(34)

    Do Until Inet.StillExecuting = False
        DoEvents
    Loop

    GetValue = Inet.GetHeader("Content-length")
    
    Do Until Inet.StillExecuting = False
        DoEvents
    Loop
    
    If IsNumeric(GetValue) = True Then
        GetSize = CLng(GetValue)
    Else
        GetSize = -1
    End If

    If GetSize <= 0 Then GetSize = -1

    m_GettingFileSize = False
    GetHTTPFileSize = GetSize
Exit Function

ErrorHandler:
    m_GettingFileSize = False
    GetHTTPFileSize = -1
End Function

Public Function CheckIP(IPToCheck As String) As Boolean

  Dim TempValues
  Dim iLoop As Long
  Dim TempByte As Byte
  
  On Error GoTo CheckIPError
  
  TempValues = Split(IPToCheck, ".")
  
  If UBound(TempValues) < 3 Then
    Exit Function
  End If
  
  For iLoop = LBound(TempValues) To UBound(TempValues)
    TempByte = TempValues(iLoop)
  Next iLoop
  CheckIP = True
  
CheckIPError:

End Function


