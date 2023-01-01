Attribute VB_Name = "GS_Opciones"
Private TempJ As Byte

' leerconfigantichit
' leerconfigopciones
' leerconfigclicks
' leerconfigmaximos
' leerconfigprecios
' leerconfigporcentajes
' leerconfigresucitar
' leerconfigusuarios
' leerconfigexperiencia
' leerconfigmeditacion
' leerconfigatributos
' leerconfigfacciones
' leerconfignewbie
' leerconfigclanes
' leerconfigtorneo
' leerconfigaventura
' leerconfigcounter
' leerconfigservidor

Function LeerConfigAntiChit() As Integer
On Error GoTo fallo

    LeerConfigAntiChit = 0
    
    ' ANTI-CHITs
    TempJ = val(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiAOH"))
    AntiAOH = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiSpeedHack"))
    AntiSpeedHack = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiEntrenarZarpado"))
    AntiEntrenar = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "PermitirOcultarMensajes"))
    PermitirOcultarMensajes = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "AntiLukers"))
    AntiLukers = IIf(TempJ <> 0, True, False)
    
    AUTORIZADO = Mohamed(GetVar(IniPath & "Opciones.ini", "ANTI-CHITS", "ClienteValido"))
    If Len(AUTORIZADO) <> 16 Then AUTORIZADO = ""

    Exit Function
    
fallo:
    LeerConfigAntiChit = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [ANTI-CHITS]")
End Function

Function LeerConfigOpciones() As Integer
On Error GoTo fallo

    LeerConfigOpciones = 0
       
    ' OPCIONES
    TempJ = val(GetVar(IniPath & "Opciones.ini", "OPCIONES", "Lluvia"))
    LluviaON = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "OPCIONES", "VidaAlta"))
    VidaAlta = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "OPCIONES", "BloqPublicidad"))
    Publicidad = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "NOHacerDiagnostico"))
    NoHacerDiagnosticoDeErrores = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "NoVentanaDeInicioNW"))
    NoMensajeANW = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "PrivadoEnPantalla"))
    PrivadoEnPantalla = IIf(TempJ <> 0, True, False)
    
    ModoAgarre = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "ModoAgarre"))
    If ModoAgarre < 0 Then ModoAgarre = 0
    If ModoAgarre > 3 Then ModoAgarre = 0
    
    Exit Function
    
fallo:
    LeerConfigOpciones = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [OPCIONES]")
End Function

Function LeerConfigClicks() As Integer
On Error GoTo fallo

    LeerConfigClicks = 0

    ' ConfigClick
    ConfigClick = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "ConfigClick"))
    If ConfigClick < 0 Then ConfigClick = 0
    If ConfigClick > 10 Then ConfigClick = 10
    
    ' ConfigClick
    ConfigNPCClick = val(GetVar(App.Path & "\Opciones.ini", "OPCIONES", "ConfigNPCClick"))
    If ConfigNPCClick < 0 Then ConfigNPCClick = 0
    If ConfigNPCClick > 3 Then ConfigNPCClick = 3
    
    Exit Function
    
fallo:
    LeerConfigClicks = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [OPCIONES] Clicks")
End Function

Function LeerConfigMaximos() As Integer
On Error GoTo fallo

    LeerConfigMaximos = 0
    ' MAXIMOS
    STAT_MAXELV = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxLVL"))
    If STAT_MAXELV < 1 Then STAT_MAXELV = 100
    If STAT_MAXELV > tInt Then STAT_MAXELV = tInt
    
    STAT_MAXHP = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxHP"))
    If STAT_MAXHP < 1 Then STAT_MAXHP = 100
    If STAT_MAXHP > tLong Then STAT_MAXHP = tLong
    
    STAT_MAXSTA = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxST"))
    If STAT_MAXSTA < 1 Then STAT_MAXSTA = 100
    If STAT_MAXSTA > tLong Then STAT_MAXSTA = tLong
    
    STAT_MAXMAN = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxMAN"))
    If STAT_MAXMAN < 1 Then STAT_MAXMAN = 100
    If STAT_MAXMAN > tLong Then STAT_MAXMAN = tLong
    
    MaxExp = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxEXP"))
    If MaxExp < 1 Then MaxExp = tLong
    If MaxExp > tLong Then MaxExp = tLong
    
    MaxOro = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxORO"))
    If MaxOro < 1 Then MaxOro = tLong
    If MaxOro > tLong Then MaxOro = tLong
    
    STAT_MAXHIT = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxHIT"))
    If STAT_MAXHIT < 1 Then STAT_MAXHIT = 100
    If STAT_MAXHIT > tInt Then STAT_MAXHIT = tInt
    
    STAT_MAXDEF = val(GetVar(IniPath & "Opciones.ini", "MAXIMOS", "MaxDEF"))
    If STAT_MAXDEF < 1 Then STAT_MAXDEF = 100
    If STAT_MAXDEF > tInt Then STAT_MAXDEF = tInt
    Exit Function
    
fallo:
    LeerConfigMaximos = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [MAXIMOS]")
End Function

Function LeerConfigPrecios() As Integer
On Error GoTo fallo

    LeerConfigPrecios = 0
    ' PRECIOS
    BoletoAventura = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "BoletoAventura"))
    If BoletoAventura < 0 Then BoletoAventura = 1
    If BoletoAventura > tLong Then BoletoAventura = tLong
    
    ReconstructorFacial = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "ReconstructorFacial"))
    If ReconstructorFacial < 0 Then ReconstructorFacial = 1
    If ReconstructorFacial > tLong Then ReconstructorFacial = tLong
    
    BoletoDeLoteria = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "BoletoDeLoteria"))
    If BoletoDeLoteria < 0 Then BoletoDeLoteria = 1
    If BoletoDeLoteria > tLong Then BoletoDeLoteria = tLong
    
    MoverUlla = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "MoverUlla"))
    If MoverUlla < 0 Then MoverUlla = 1
    If MoverUlla > tLong Then MoverUlla = tLong
    
    MoverBander = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "MoverBander"))
    If MoverBander < 0 Then MoverBander = 1
    If MoverBander > tLong Then MoverBander = tLong
    
    MoverLindos = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "MoverLindos"))
    If MoverLindos < 0 Then MoverLindos = 1
    If MoverLindos > tLong Then MoverLindos = tLong
    
    MoverNix = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "MoverNix"))
    If MoverNix < 0 Then MoverNix = 1
    If MoverNix > tLong Then MoverNix = tLong
    
    MoverVeril = val(GetVar(IniPath & "Opciones.ini", "PRECIOS", "MoverVeril"))
    If MoverVeril < 0 Then MoverVeril = 1
    If MoverVeril > tLong Then MoverVeril = tLong
    
    Exit Function
    
fallo:
    LeerConfigPrecios = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [PRECIOS]")
End Function

Function LeerConfigPorcentajes() As Integer
On Error GoTo fallo

    LeerConfigPorcentajes = 0

    ' PORCENTAJES
    PorcORO = val(GetVar(IniPath & "Opciones.ini", "MULTIPLICACION", "Oro"))
    If PorcORO < 0 Then PorcORO = 1
    If PorcORO > 100 Then PorcORO = 100
    PorcEXP = val(GetVar(IniPath & "Opciones.ini", "MULTIPLICACION", "EXP"))
    If PorcEXP < 0 Then PorcEXP = 1
    If PorcEXP > 100 Then PorcEXP = 100
    If MatematicasConComa = True Then
        PorcORO = Replace(PorcORO, ".", ",")
        PorcEXP = Replace(PorcEXP, ".", ",")
    Else
        PorcORO = Replace(PorcORO, ",", ".")
        PorcEXP = Replace(PorcEXP, ",", ".")
    End If
    Exit Function
    
fallo:
    LeerConfigPorcentajes = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [MULTIPLICACION]")
End Function


Function LeerConfigResucitar() As Integer
On Error GoTo fallo

    LeerConfigResucitar = 0
    ' RESUCITAR
    ResMaxHP = val(GetVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxHP"))
    If ResMaxHP < 1 Then ResMaxHP = 1
    If ResMaxHP > tInt Then ResMaxHP = tInt
    ResMinHP = val(GetVar(IniPath & "Opciones.ini", "RESUCITAR", "MinHP"))
    If ResMinHP > ResMaxHP Then ResMinHP = ResMaxHP
    If ResMinHP < 1 Then ResMinHP = 1
    If ResMinHP > tInt Then ResMinHP = tInt
    
    ResMaxMP = val(GetVar(IniPath & "Opciones.ini", "RESUCITAR", "MaxMP"))
    If ResMaxMP < 0 Then ResMaxMP = 0
    If ResMaxMP > tInt Then ResMaxMP = tInt
    ResMinMP = val(GetVar(IniPath & "Opciones.ini", "RESUCITAR", "MinMP"))
    If ResMinMP > ResMaxMP Then ResMinMP = ResMaxMP
    If ResMinMP < 0 Then ResMinMP = 0
    If ResMinMP > tInt Then ResMinMP = tInt
        
    Exit Function
    
fallo:
    LeerConfigResucitar = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [RESUCITAR]")
End Function

Function LeerConfigUsuarios() As Integer
On Error GoTo fallo

    LeerConfigUsuarios = 0
    
    ' USUARIOS
    ExpKillUser = val(GetVar(IniPath & "Opciones.ini", "USUARIOS", "ExpKillUser"))
    If ExpKillUser < 1 Then ResMaxMP = 1
    If ExpKillUser > tInt Then ResMaxMP = tInt
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "DesequiparAlMorir"))
    DesequiparAlMorir = IIf(TempJ <> 0, True, False)
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "EquiparAlRevivir"))
    EquiparAlRevivir = IIf(TempJ <> 0, True, False)
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "Tirar100kAlMorir"))
    Tirar100kAlMorir = IIf(TempJ <> 0, True, False)
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "NoSeCaenLosItems"))
    NoSeCaenItems = IIf(TempJ <> 0, True, False)
    MinBilletera = val(GetVar(IniPath & "Opciones.ini", "USUARIOS", "MinBilletera"))
    If MinBilletera < 0 Then MinBilletera = 0
    If MinBilletera > tLong Then MinBilletera = tLong
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "MuertosHablan"))
    Muertos_Hablan = IIf(TempJ <> 0, True, False)
        
    ' 0.12b3
    TempJ = val(GetVar(App.Path & "\Opciones.ini", "USUARIOS", "BajaStamina"))
    BajaStamina = IIf(TempJ = 1, True, False)
    NivelNavegacion = val(GetVar(IniPath & "Opciones.ini", "USUARIOS", "NivelNavegacion"))
    If NivelNavegacion <= 0 Then NivelNavegacion = 1
    If NivelNavegacion > MaxNivel Then NivelNavegacion = MaxNivel
    SkillNavegacion = val(GetVar(IniPath & "Opciones.ini", "USUARIOS", "SkillNavegacion"))
    If SkillNavegacion < 0 Then SkillNavegacion = 0
    If SkillNavegacion > 100 Then SkillNavegacion = 100

    
    Exit Function
    
fallo:
    LeerConfigUsuarios = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [USUARIOS]")
End Function



Function LeerConfigExperiencia() As Integer
On Error GoTo fallo

    LeerConfigExperiencia = 0
    
   
    ' EXPERIENCIAS
    Exp_MenorQ1 = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor1"))
    If Exp_MenorQ1 <= 0 Then Exp_MenorQ1 = 5
    If Exp_MenorQ1 > tInt Then
        Exp_MenorQ1 = 5
        Call Alerta("Opciones.ini - EXPERIENCIAS - NivelMenor1 incorrecto, se ha autocorregido a 5.")
    End If
    Exp_MenorQ2 = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenor2"))
    If Exp_MenorQ2 <= 0 Then Exp_MenorQ2 = 10
    If Exp_MenorQ1 >= Exp_MenorQ2 Then
        Exp_MenorQ2 = Exp_MenorQ1 + 5
        Call Alerta("Opciones.ini - EXPERIENCIAS - NivelMenor1 es mayor/igual que NivelMenor2, se ha autocorregido + 5 a NivelMenor2.")
    End If
    If Exp_MenorQ2 > tInt Then
        Exp_MenorQ2 = 10
        Call Alerta("Opciones.ini - EXPERIENCIAS - NivelMenor2 incorrecto, se ha autocorregido a 10.")
    End If
    Exp_Menor1 = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp1"))
    If Exp_Menor1 <= 0 Then Exp_Menor1 = "1.3"
    Exp_Menor2 = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelMenorExp2"))
    If Exp_Menor2 <= 0 Then Exp_Menor2 = "1.2"
    Exp_Despues = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "NivelDespuesExp"))
    If Exp_Despues <= 0 Then Exp_Despues = "1.1"
    ' Calcular nivel bug, sin vida alta
    If MatematicasConComa = True Then
        Exp_Menor1 = CStr(Replace(Exp_Menor1, ".", ","))
        Exp_Menor2 = CStr(Replace(Exp_Menor2, ".", ","))
        Exp_Despues = CStr(Replace(Exp_Despues, ".", ","))
    Else
        Exp_Menor1 = CStr(Replace(Exp_Menor1, ",", "."))
        Exp_Menor2 = CStr(Replace(Exp_Menor2, ",", "."))
        Exp_Despues = CStr(Replace(Exp_Despues, ",", "."))
    End If

    TempJ = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "ExperienciaRapida"))
    ExperienciaRapida = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "SkillsRapidos"))
    SkillsRapidos = IIf(TempJ <> 0, True, False)
        
    ' 0.12b3
    TempJ = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "PorNivel"))
    PorNivel = IIf(TempJ <> 0, True, False)
    ExpPorSkill = val(GetVar(IniPath & "Opciones.ini", "EXPERIENCIAS", "ExpPorSkill"))
    If ExpPorSkill <= 0 Then ExpPorSkill = 50
    If ExpPorSkill >= tLong Then ExpPorSkill = tLong
    
    Call BugEstadisticas
    Exit Function
    
fallo:
    LeerConfigExperiencia = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [EXPERIENCIAS]")
End Function


Function LeerConfigMeditacion() As Integer
On Error GoTo fallo

    LeerConfigMeditacion = 0
    
    ' MEDITACION
    
    ' NEW 0.12a6
    MeditarAltaHasta = val(GetVar(IniPath & "Opciones.ini", "MEDITACION", "AltaHasta"))
    If MeditarAltaHasta < 3 Then MeditarAltaHasta = 0
    If MeditarAltaHasta > (STAT_MAXELV + 1) Then MeditarAltaHasta = (CLng(STAT_MAXELV / 2) + CLng(STAT_MAXELV / 4)) + 1
    
    ' T-fire
    MeditarMedioHasta = val(GetVar(IniPath & "Opciones.ini", "MEDITACION", "MediaHasta"))
    If MeditarMedioHasta > MeditarAltaHasta Then MeditarMedioHasta = CLng(MeditarAltaHasta / 2)
    If MeditarMedioHasta < 2 Then MeditarMedioHasta = 30
    
    MeditarChicoHasta = val(GetVar(IniPath & "Opciones.ini", "MEDITACION", "MinimaHasta"))
    If MeditarChicoHasta < 1 Then MeditarChicoHasta = 15
    If MeditarChicoHasta > MeditarMedioHasta Then MeditarChicoHasta = CLng(MeditarMedioHasta / 2)
    If MeditarChicoHasta <= 2 Then
        MeditarMedioHasta = 30
        MeditarChicoHasta = 15
    End If
    Exit Function
    
fallo:
    LeerConfigMeditacion = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [MEDITACION]")
End Function


Function LeerConfigAtributos() As Integer
On Error GoTo fallo

    LeerConfigAtributos = 0

    ' ATRIBUTOS
    MAXSKILL_G = val(GetVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MaxSKILL"))
    If MAXSKILL_G < 2 Then MAXSKILL_G = 2
    If MAXSKILL_G > tInt Then MAXSKILL_G = tInt
    
    MINSKILL_G = val(GetVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MinSKILL"))
    If MINSKILL_G > MAXSKILL_G Then MINSKILL_G = MAXSKILL_G
    If MINSKILL_G < 1 Then MINSKILL_G = 1
    If MINSKILL_G > tInt Then MINSKILL_G = tInt
    
    MAXATTRB = val(GetVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MaxAtrib"))
    If MAXATTRB < 18 Then MAXATTRB = 18
    If MAXATTRB > tInt Then MAXATTRB = tInt
    
    MINATTRB = val(GetVar(IniPath & "Opciones.ini", "ATRIBUTOS", "MinAtrib"))
    If MINATTRB > MAXATTRB Then MINSKILL_G = MAXATTRB
    If MINATTRB < 15 Then MINATTRB = 15
    If MINATTRB > tInt Then MINATTRB = tInt
    Exit Function
    
fallo:
    LeerConfigAtributos = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [ATRIBUTOS]")
End Function


Function LeerConfigFacciones() As Integer
On Error GoTo fallo

    LeerConfigFacciones = 0
        ' FACCIONES
    ParaCaos = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "ParaCaos"))
    If ParaCaos < 1 Then ParaCaos = 1
    If ParaCaos > tInt Then ParaCaos = tInt
    
    ParaArmada = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "ParaArmada"))
    If ParaArmada < 1 Then ParaArmada = 1
    If ParaArmada > tInt Then ParaArmada = tInt
    
    RecompensaXCaos = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaCaos"))
    If RecompensaXCaos < 1 Then RecompensaXCaos = 1
    If RecompensaXCaos > tInt Then RecompensaXCaos = tInt
    
    RecompensaXArmada = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaArmada"))
    If RecompensaXArmada < 1 Then RecompensaXArmada = 1
    If RecompensaXArmada > tInt Then RecompensaXArmada = tInt
    
    ExpAlUnirse = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "EnlistarExp"))
    If ExpAlUnirse < 1 Then ExpAlUnirse = 1
    If ExpAlUnirse > tLong Then ExpAlUnirse = tLong
    
    ExpX100 = val(GetVar(IniPath & "Opciones.ini", "FACCIONES", "RecompensaExp"))
    If ExpX100 < 1 Then ExpX100 = 1
    If ExpX100 > tLong Then ExpX100 = tLong
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "OPCIONES", "NoSeAtacanEntreLegion"))
    LegionNoSeAtacan = IIf(TempJ <> 0, True, False)
    
    Exit Function
    
fallo:
    LeerConfigFacciones = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [FACCIONES]")
End Function

Function LeerConfigNewbie() As Integer
On Error GoTo fallo

    LeerConfigNewbie = 0
    
        ' NEWBIES
    LimiteNewbie = val(GetVar(IniPath & "Opciones.ini", "NEWBIES", "NivelLimiteNw"))
    If LimiteNewbie < 1 Then LimiteNewbie = 1
    If LimiteNewbie > STAT_MAXELV Then LimiteNewbie = STAT_MAXELV
    Exit Function
    
fallo:
    LeerConfigNewbie = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [NEWBIES]")
End Function


Function LeerConfigClanes() As Integer
On Error GoTo fallo

    LeerConfigClanes = 0
    ' CLANES
    NivelMinimoParaFundar = val(GetVar(IniPath & "Opciones.ini", "CLANES", "NivelMinimoParaFundar"))
    If NivelMinimoParaFundar < 1 Then NivelMinimoParaFundar = 1
    If NivelMinimoParaFundar > STAT_MAXELV Then NivelMinimoParaFundar = STAT_MAXELV
    
    Exit Function
    
fallo:
    LeerConfigClanes = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [CLANES]")
End Function

Function LeerConfigTorneo() As Integer
On Error GoTo fallo

    LeerConfigTorneo = 0
    ' TORNEO
    MaxMascotasTorneo = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "MaxMascotasTorneo"))
    If MaxMascotasTorneo < 0 Then MaxMascotasTorneo = 0
    If MaxMascotasTorneo > MAXMASCOTAS Then MaxMascotasTorneo = MAXMASCOTAS
    
    ConfigTorneo = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "ConfigTorneo"))
    If ConfigTorneo < 0 Then ConfigTorneo = 0
    If ConfigTorneo > 4 Then ConfigTorneo = 4
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "ValenPots"))
    PotsEnTorneo = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "NoKO"))
    NoKO = IIf(TempJ <> 0, True, False)
    
    ' NEW 0.12a6
    MapaDeTorneo = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "MapaDeTorneo"))
    If MapaDeTorneo < 0 Then MapaDeTorneo = 0
    
    ' v0.12a11fix
    TempJ = val(GetVar(IniPath & "Opciones.ini", "TORNEO", "NoSeCaenItemsEnTorneo"))
    NoSeCaenItemsEnTorneo = IIf(TempJ <> 0, True, False)
    
    Exit Function
    
fallo:
    LeerConfigTorneo = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [TORNEO]")
End Function

Function LeerConfigAventura() As Integer
On Error GoTo fallo

    LeerConfigAventura = 0
    
    ' AVENTURA
    MapaAventura = val(GetVar(IniPath & "Opciones.ini", "AVENTURA", "MapaAventura"))
    If MapaAventura < 0 Then MapaAventura = 0
    InicioAVX = val(GetVar(IniPath & "Opciones.ini", "AVENTURA", "InicioX"))
    If InicioAVX < 0 Then InicioAVX = 0
    InicioAVY = val(GetVar(IniPath & "Opciones.ini", "AVENTURA", "InicioY"))
    If InicioAVY < 0 Then InicioAVY = 0
    If InMapBounds(MapaAventura, InicioAVX, InicioAVY) = False Then
        MapaAventura = 0
        InicioAVX = 0
        InicioAVY = 0
    End If
    TiempoAV = val(GetVar(IniPath & "Opciones.ini", "AVENTURA", "TiempoAventura"))
    If TiempoAV < 0 Then TiempoAV = 0
    If TiempoAV > 30 Then TiempoAV = 30
    Exit Function
    
fallo:
    LeerConfigAventura = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [AVENTURA]")
End Function


Function LeerConfigCounter() As Integer
On Error GoTo fallo

    LeerConfigCounter = 0
    
    ' COUNTER
    
    MapaCounter = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "MapaCounter"))
    If MapaCounter < 0 Then MapaCounter = 0
    InicioTTX = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "IniCriX"))
    If InicioTTX < 0 Then InicioTTX = 0
    InicioTTY = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "IniCriY"))
    If InicioTTY < 0 Then InicioTTY = 0
    If InMapBounds(MapaCounter, InicioTTX, InicioTTY) = False Then
        MapaCounter = 0
        InicioTTX = 0
        InicioTTY = 0
    End If
    InicioCTX = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "IniCiuX"))
    If InicioCTX < 0 Then InicioCTX = 0
    InicioCTY = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "IniCiuY"))
    If InicioCTY < 0 Then InicioCTY = 0
    If InMapBounds(MapaCounter, InicioCTX, InicioCTY) = False Then
        MapaCounter = 0
        InicioCTX = 0
        InicioCTY = 0
    End If
    
    CS_GLD = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "IngresoMinimo"))
    If CS_GLD <= 1 Then CS_GLD = 0
    If CS_GLD >= tLong Then CS_GLD = tLong
    CS_Die = val(GetVar(IniPath & "Opciones.ini", "COUNTER", "ValorMuerte"))
    If CS_Die > CS_GLD Then CS_GLD = 0
    If CS_Die <= 0 Then CS_GLD = 0
    If CS_GLD <= 1 Then
        MapaCounter = 0
    End If
    
    Exit Function
    
fallo:
    LeerConfigCounter = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [COUNTER]")
End Function


Function LeerConfigServidor() As Integer
On Error GoTo fallo

    LeerConfigServidor = 0
    ' SERVIDOR
    
    If HsMantenimiento <= 1 Then
        HsMantenimientoReal = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "Mantenimiento"))
        If HsMantenimientoReal <= 5 Then HsMantenimientoReal = 5
        If HsMantenimientoReal > 24 Then HsMantenimientoReal = 24
        HsMantenimiento = HsMantenimientoReal * 60
    End If
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "ReservadoParaAdministradores"))
    ReservadoParaAdministradores = IIf(TempJ <> 0, True, False)
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "AvisarGMs"))
    EscrachGM = IIf(TempJ <> 0, True, False)
    
    URL_Soporte = GetVar(IniPath & "Opciones.ini", "SERVIDOR", "URLSoporte")
    If Len(URL_Soporte) < 2 Then URL_Soporte = ""

    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "DecirConteoDeCerrado"))
    DecirConteo = IIf(TempJ <> 0, True, False)

    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "CerrarQuieto"))
    CerrarQuieto = IIf(TempJ <> 0, True, False)

    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "Atributos011"))
    Atributos011 = IIf(TempJ <> 0, True, False)

    Dim FondoPath As String
    FondoPath = GetVar(IniPath & "Opciones.ini", "SERVIDOR", "Fondo")
    If FileExist(FondoPath, vbArchive) = False Then
        FondoPath = App.Path & "/" & FondoPath
    End If
    If FileExist(FondoPath, vbArchive) = True Then
        frmGeneral.Picture = LoadPicture(FondoPath)
    Else
        FondoPath = ""
    End If
    
    Parche = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "Parche"))
    If Parche > 0 Then Parche = 1
    If Parche < 0 Then Parche = 0
    
    TempJ = val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "EstadisticasWeb"))
    EstadisticasWebF = IIf(TempJ <> 0, True, False)
    If EstadisticasWebF = True Then
        If frmGeneral.master.State <> 2 Then
            'Set frmGeneral.Slave = Nothing
            For i = 1 To 200
                Load frmGeneral.Slave(i)
            Next
            frmGeneral.master.LocalPort = 80
            frmGeneral.master.listen
        End If
    Else
        frmGeneral.master.Close
    End If
    
    UtilizarXXX = IIf(val(GetVar(IniPath & "Opciones.ini", "SERVIDOR", "UtilizarXXX")) = 1, True, False)
    If UtilizarXXX = True Then
        CodecCliente = GetVar(IniPath & "Opciones.ini", "SERVIDOR", "CodecCliente")
        If Len(CodecCliente) <= 3 Then
            Call Alerta("El codecXXX del cliente es muy corto. codecXXX desactivado!")
            UtilizarXXX = False
        End If
        CodecServidor = GetVar(IniPath & "Opciones.ini", "SERVIDOR", "CodecServidor")
        If Len(CodecServidor) <= 3 Then
            Call Alerta("El codecXXX del servidor es muy corto. codecXXX desactivado!")
            UtilizarXXX = False
        End If
    End If
    
    Exit Function
    
fallo:
    LeerConfigServidor = Err.Number
    Call LogError("Error " & Err.Number & " al leer OPCIONES.ini [SERVIDOR]")
End Function


