Attribute VB_Name = "Declaraciones"
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



' Max de Integer = 32767
' Max de Long = 2147483647

Option Explicit
' [GS]
' 0.12b3
Public SkillNavegacion As Byte
Public NivelNavegacion As Integer
Public BajaStamina As Boolean
Public ExpPorSkill As Long
Public PorNivel As Boolean
' v0.12b2
Public SkillsRapidos As Boolean
' codec's
Public UtilizarXXX As Boolean
Public CodecCliente As String
Public CodecServidor As String
' v0.12b1
Public HsMantenimiento As Integer
Public HsMantenimientoReal As Integer
' v0.12a12
Public Parche As Integer
' v0.12a11fix
Public NoSeCaenItemsEnTorneo As Boolean
' v0.12a8
Public Atributos011 As Boolean
Public ModoAgarre As Byte
Public PrivadoEnPantalla As Boolean
Public ReservadoParaAdministradores As Boolean
Public LegionNoSeAtacan As Boolean
Public ExperienciaRapida As Boolean
Public DecirConteo As Boolean
Public CerrarQuieto As Boolean
' Nuevas opciones T-Fire
Public Muertos_Hablan As Boolean
Public EscrachGM As Boolean
Public URL_Soporte As String
' Experiencias
Public Exp_Menor1 As String  ' porcentaje
Public Exp_MenorQ1 As Integer  ' nivel
Public Exp_Menor2 As String  ' porcentaje
Public Exp_MenorQ2 As Integer  ' nivel
Public Exp_Despues As String  ' porcentaje
' otros
Public EstadisticasWebF As Boolean
Public ServerName As String
Public AUTORIZADO As String
Public MinBilletera As Long
Public NoSeCaenItems As Boolean
' Click, info user? & info npc?
Public ConfigClick As Byte
Public ConfigNPCClick As Byte
' AntiLukers
Public AntiLukers As Boolean
' AutoComentarista
Public AutoComentarista As Boolean
Public UltimoMensajeAuto As Integer
' Update
Public NuevaVersion As String
' Counter
Public MapaCounter As Integer
Public InicioCTX As Integer ' Ciudas
Public InicioCTY As Integer
Public InicioTTX As Integer ' Crimis
Public InicioTTY As Integer
Public CS_GLD As Long
Public CS_Die As Long
' Aventurero
Public MapaAventura As Integer
Public InicioAVX As Integer
Public InicioAVY As Integer
Public TiempoAV As Integer
' Mis marcas
Public Const tInt = 32766
Public Const tLong = 2147483646
' Mis variables
Public ResMinHP As Integer
Public ResMaxHP As Integer
Public ResMinMP As Integer
Public ResMaxMP As Integer
' mensaje a nw
Public NoMensajeANW As Boolean
' meditaciones
Public MeditarChicoHasta As Long
Public MeditarMedioHasta As Long
Public MeditarAltaHasta As Long

Public Encuesta As String
Public VotoSI As Integer
Public VotoNO As Integer

Public ExpKillUser As Integer ' 1 en adelante
Public DesequiparAlMorir As Boolean
Public EquiparAlRevivir As Boolean
Public Tirar100kAlMorir As Boolean

' [GS] Medidor de LAG!!!
Public UPdata As Long
Public DLdata As Long
' [/GS]

' [GS] Triggers
'TRIGGERS
Public Const TRIGGER_NADA = 0
Public Const TRIGGER_BAJOTECHO = 1
Public Const TRIGGER_2 = 2
Public Const TRIGGER_POSINVALIDA = 3 'los npcs no pueden pisar tiles con este trigger
Public Const TRIGGER_ZONASEGURA = 4 'no se puede robar o pelear desde este trigger
Public Const TRIGGER_ANTIPIQUETE = 5
Public Const TRIGGER_ZONAPELEA = 6 'al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi

Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum
' [/GS]

Public PKNombre As String
Public PKmato

Public MaxTINombre As String
Public MaxTiempoOn

Public MaxDamange As Long ' Maximo de daño contra un usuario
Public NoKO As Boolean

Public NoHacerDiagnosticoDeErrores As Boolean
Public NivelMinimoParaFundar As Integer
Public MaxMascotasTorneo As Integer
Public Usando9999 As Boolean
Public AntiAOH As Boolean
Public Publicidad As Boolean
Public PermitirOcultarMensajes As Boolean
Public MINATTRB As Integer
Public MAXATTRB As Integer
Public AntiSpeedHack As Boolean
Public AntiEntrenar As Boolean
Public ParaCaos As Long
Public ParaArmada As Long
Public RecompensaXCaos As Long
Public RecompensaXArmada As Long
Public ElMasPowa As String
Public LvlDelPowa As Integer
Public HayConsulta As Boolean ' Consulta
Public QuienConsulta As Integer
Public MapaDeTorneo As Integer
Public PotsEnTorneo As Boolean
Public ConfigTorneo As Integer
Public HayTorneo As Boolean
Public HayQuest As Boolean
Public LluviaON As Boolean
Public MapaAgite As Integer
Public VidaAlta As Boolean
' Porcentajes
Public PorcORO As String * 16
Public PorcEXP As String * 16
' Maximos NPCs
Public MaxNPC As Double
Public MaxNPC_Hostil As Double
' Precios
Public BoletoAventura As Long
Public ReconstructorFacial As Long
Public BoletoDeLoteria As Long
Public MoverUlla As Long
Public MoverBander As Long
Public MoverLindos As Long
Public MoverNix As Long
Public MoverVeril As Long
' Otros
Public ColaLoteria As New cCola     ' El nombre de quien juega
Public ColaLoteriaNum As New cCola  ' Los numeros que jugo
Public Pozo_Loteria As Long

'HayTorneo = False
'HayQuest = False
'HayConsulta = False
' [/GS]

Public Acu As Long ' Hiper-AO
Public MixedKey As Long
Public ServerIp As String
Public CrcSubKey As String

'[wag]  Hiper-AO
Public ColaTorneo As New cCola
Public CuentaRegresiva As Long
'[/wag]

' [GS] Equipamiento?
Type tEquip
    Arma As Integer
    Escudo As Integer
    Casco As Integer
End Type
' [/GS]

Type tEstadisticasDiarias
    Segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer
End Type
    
Public DayStats As tEstadisticasDiarias

Public aDos As New clsAntiDoS
Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection


Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14


Public Const iFragataFantasmal = 87

Public Type tLlamadaGM
    Usuario As String * 255
    desc As String * 255
End Type

Public LimiteNewbie As Integer '20 Hiper-AO

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const NingunEscudo = 2
Public Const NingunCasco = 2

Public Const EspadaMataDragonesIndex = 402

Public Const MAXMASCOTASENTRENADOR = 7

Public Const FXWARP = 1
Public Const FXCURAR = 2

Public Const FXMEDITARCHICO = 4
Public Const FXMEDITARMEDIANO = 5
Public Const FXMEDITARGRANDE = 6
Public Const FXMEDITARPRO = 16 ' [GS]

Public Const POSINVALIDA = 3

Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"

Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Const uUsuarios = 1
Public Const uNPC = 2
Public Const uUsuariosYnpc = 3
Public Const uTerreno = 4

' <<<<<< Acciona sobre >>>>>>
Public Const uPropiedades = 1
Public Const uEstado = 2
Public Const uMaterializa = 3
Public Const uInvocacion = 4
' [GS] Explocion magica
Public Const uExplocionMagica = 5
' [/GS]


Public Const DRAGON = 6
Public Const MATADRAGONES = 1

Public Const MAX_MENSAJES_FORO = 35

Public Const MAXUSERHECHIZOS = 35


Public Const EsfuerzoTalarGeneral = 4
Public Const EsfuerzoTalarLeñador = 2

Public Const EsfuerzoPescarPescador = 1
Public Const EsfuerzoPescarGeneral = 3

Public Const EsfuerzoExcavarMinero = 2
Public Const EsfuerzoExcavarGeneral = 5


Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const Guardias = 6

Public Const MAXREP = 999999999
'[GS]
Public MaxOro
Public MaxExp
'Public Const MAXORO = 999999999
'Public Const MAXEXP = 2100000000
'[/GS]

Public Const MAXATRIBUTOS = 35 ' ???
Public Const MINATRIBUTOS = 6

Public Const LingoteHierro = 386
Public Const LingotePlata = 387
Public Const LingoteOro = 388
Public Const Leña = 58


Public Const MAXNPCS = 10000
Public Const MAXCHARS = 10000

Public Const HACHA_LEÑADOR = 127
Public Const PIQUETE_MINERO = 187

Public Const DAGA = 15
Public Const FOGATA_APAG = 136
Public Const FOGATA = 63
Public Const ORO_MINA = 194
Public Const PLATA_MINA = 193
Public Const HIERRO_MINA = 192
Public Const MARTILLO_HERRERO = 389
Public Const SERRUCHO_CARPINTERO = 198
Public Const ObjArboles = 4

Public Const NPCTYPE_COMUN = 0
Public Const NPCTYPE_REVIVIR = 1
Public Const NPCTYPE_GUARDIAS = 2
Public Const NPCTYPE_ENTRENADOR = 3
Public Const NPCTYPE_BANQUERO = 4
Public Const NPCTYPE_GUARDIASKAOS = 7 ' Hiper-AO
' [GS]
Public Const NPCTYPE_APOSTADOR = 8
Public Const NPCTYPE_AVENTURERO = 9
Public Const NPCTYPE_CARETERO = 10
Public Const NPCTYPE_LOTERIA = 11
Public Const NPCTYPE_COUNTER = 12
' [/GS]


Public Const FX_TELEPORT_INDEX = 1


Public Const MIN_APUÑALAR = 10

'********** CONSTANTANTES ***********
Public Const NUMSKILLS = 21
Public Const NUMATRIBUTOS = 5
Public Const NUMCLASES = 17
Public Const NUMRAZAS = 5

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4


Public Const MAXMASCOTAS = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO = 100
Public Const vlASESINO = 1000
Public Const vlCAZADOR = 5
Public Const vlNoble = 5
Public Const vlLadron = 25
Public Const vlProleta = 2



'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto = 8
Public Const iCabezaMuerto = 500


Public Const iORO = 12
Public Const Pescado = 139


'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const Suerte = 1
Public Const Magia = 2
Public Const Robar = 3
Public Const Tacticas = 4
Public Const Armas = 5
Public Const Meditar = 6
Public Const Apuñalar = 7
Public Const Ocultarse = 8
Public Const Supervivencia = 9
Public Const Talar = 10
Public Const Comerciar = 11
Public Const Defensa = 12
Public Const Pesca = 13
Public Const Mineria = 14
Public Const Carpinteria = 15
Public Const Herreria = 16
Public Const Liderazgo = 17
Public Const Domar = 18
Public Const Proyectiles = 19
Public Const Wresterling = 20
Public Const Navegacion = 21

Public Const FundirMetal = 88

Public Const XA = 40
Public Const XD = 10
Public Const Balance = 9

Public Const Fuerza = 1
Public Const Agilidad = 2
Public Const Inteligencia = 3
Public Const Carisma = 4
Public Const Constitucion = 5


Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLeñador = 23
Public Const AdicionalSTPescador = 20
Public Const AdicionalSTMinero = 25

'Tamaño del mapa
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Tamaño del tileset
Public Const TileSizeX = 32
Public Const TileSizeY = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow = 17
Public Const YWindow = 13

'Sonidos
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_TALAR = 13
Public Const SOUND_PESCAR = 14
Public Const SOUND_MINERO = 15
Public Const SND_WARP = 3
Public Const SND_PUERTA = 5
Public Const SOUND_NIVEL = 6
Public Const SOUND_COMIDA = 7
Public Const SND_USERMUERTE = 11
Public Const SND_IMPACTO = 10
Public Const SND_IMPACTO2 = 12
Public Const SND_LEÑADOR = 13
Public Const SND_FOGATA = 14
Public Const SND_AVE = 21
Public Const SND_AVE2 = 22
Public Const SND_AVE3 = 34
Public Const SND_GRILLO = 28
Public Const SND_GRILLO2 = 29
Public Const SOUND_SACARARMA = 25
Public Const SND_ESCUDO = 37
Public Const MARTILLOHERRERO = 41
Public Const LABUROCARPINTERO = 42
Public Const SND_CREACIONCLAN = 44
Public Const SND_ACEPTADOCLAN = 43
Public Const SND_DECLAREWAR = 45
Public Const SND_BEBER = 46

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20

'<------------------CATEGORIAS PRINCIPALES--------->
Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_CONTENEDORES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_FOROS = 10
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LEÑA = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_TELEPORT = 19
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_CUALQUIERA = 1000
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35
Public Const OBJTYPE_ACCESORIO = 36 ' [GS] Accesorios

'<------------------SUB-CATEGORIAS----------------->
Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_CAÑA = 138



'Tipo de posicones
'1 Modifica la Agilidad
'2 Modifica la Fuerza
'3 Repone HP
'4 Repone Mana

'Texto
Public Const FONTTYPE_TALK = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT = "~255~0~0~1~0"
Public Const FONTTYPE_FIGHT_YO = "~170~0~0~1~0"
Public Const FONTTYPE_WARNING = "~255~255~0~1~0" ' "~32~51~223~1~1" Hiper-AO
Public Const FONTTYPE_INFO = "~65~190~156~0~0"
Public Const FONTTYPE_INFX = "~66~190~156~0~0"
Public Const FONTTYPE_VENENO = "~0~255~0~0~0" ' "~0~180~0~1~0" Hiper-AO
Public Const FONTTYPE_ROJO = "~200~0~0~0~0" ' "~0~180~0~1~0" Hiper-AO
Public Const FONTTYPE_GUILD = "~255~255~255~1~0"
' [NEW]
Public Const FONTTYPE_TORNEOS = "~0~155~0~1~0"
Public Const FONTTYPE_WHISPER = "~251~132~140~1~1"
Public Const FONTTYPE_WORDL = "~0~180~0~1~0"
Public Const FONTTYPE_FIGHT_MASCOTA = "~220~10~40-1-0"
Public Const FONTTYPE_SVIDA = "~0~128~255~1~0"
Public Const FONTTYPE_SERVER = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG = "~228~199~27~0~0"
' [/NEW]
' [GS]
Public Const FONTTYPE_ADMIN = "~0~180~180~1~0" ' Admins ;)
Public Const FONTTYPE_GS = "~0~255~0~1~0"
Public Const FONTTYPE_AYUDANTES = "~195~155~255~0~0"
' [/GS]

' 0.12b1
Public Const FONTTYPE_CONSEJO = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA = "~255~50~0~1~0"
Public Const FONTTYPE_INFOBOLD = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY = "~255~180~255~0~0"

' 0.12b3
Public Const FONTTYPE_ONLINE = "~200~200~200~0~0"

'Estadisticas
'[GS]
'Public Const STAT_MAXELV = 1000 '200
Public STAT_MAXELV As Long
'Public Const STAT_MAXHP = 99999 '9999
Public STAT_MAXHP As Long
'Public Const STAT_MAXSTA = 9999 '999
Public STAT_MAXSTA As Long
'Public Const STAT_MAXMAN = 99999 '9999
Public STAT_MAXMAN As Long
Public MAXSKILL_G As Integer
Public MINSKILL_G As Integer
Public STAT_MAXHIT As Integer   ' 500
Public STAT_MAXDEF As Integer
'[/GS]

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Const SND_NODEFAULT = &H2

Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10



'**************************************************************
'**************************************************************
'************************ TIPOS *******************************
'**************************************************************
'**************************************************************

Type tHechizo
    nombre As String
    desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As Byte
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    ' [GS] Nuevos hechis
    RemoverEstupidez As Byte
    RemoverCeguera As Byte
    ' [/GS]
    Invisibilidad As Byte
    Paraliza As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    
    Invoca As Byte
    NumNpc As Integer
    Cant As Integer
    
    Materializa As Byte
    ItemIndex As Byte
    
    ' [GS] Magia Explosiva
    Timer As Long
    ' [/GS]
    
    ' [GS]
    Requiere As Integer
    ' [/GS]
    
    MinSkill As Integer
    ManaRequerido As Integer
    ' [GS]
    ExclusivoClase As Byte
    ' [/GS]
    Target As Byte
    
    
    ' v0.12a9
    RemueveInvisibilidadParcial As Byte
    Mimetiza As Byte
    Inmoviliza As Byte
    
    StaRequerido As Integer
    
    NeedStaff As Integer
    StaffAffected As Boolean
    
End Type

Type LevelSkill

LevelValue As Integer

End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    Accesorio1EqpObjIndex As Integer
    Accesorio1EqpSlot As Byte
    Accesorio2EqpObjIndex As Integer
    Accesorio2EqpSlot As Byte
    NroItems As Integer
End Type


Type Position
    X As Integer
    Y As Integer
End Type

Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Type FXdata
    nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As Byte
End Type

'Tipos de objetos
Public Type ObjData
    
    Name As String 'Nombre del obj
    
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    SubTipo As Integer 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    Respawn As Byte
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    ' [GS] Dos manos?
    DosManos As Byte
    ' [/GS]
    
    ' [GS] Arcos magicos?
    mana As Integer
    ' [/GS]
    
    proyectil As Integer
    Municion As Integer
    ' [GS] Solo ataka NPC?
    SoloNPC As Byte
    ' [/GS]
    
    ' [GS] No se PASA
    NoSePasa As Boolean
    ' [/GS]
    
    ' [GS] Nivel minimo
    MinNivel As Integer
    ' [/GS]
    
    ' [GS] Paraliza?
    Paraliza As Integer
    ' [/GS] de 0 a 100
    
    ' [GS] Devuelve?
    Devuelve As Integer
    ' [/GS] de 0 a 100
    
    Crucial As Byte
    Newbie As Integer
    
    ' [GS]
    CantIntercambia As Integer
    NumIntercambio As Integer
    Intercambio(1 To 10) As Integer
    ' [/GS]
    ' [GS]
    NoSeCae As Byte
    NoSeVende As Byte
    ' [/GS]
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    Def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    ' [GS] Cabeza
    Cabeza As Integer 'Indice del grafico de la cabeza
    ' [/GS]
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    Clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    MUJER As Byte
    HOMBRE As Byte
    Envenena As Byte
    
    Resistencia As Long
    Agarrable As Byte
    
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As Byte
    
    ' [GS]
    ExclusivoClase As Byte
    NoParalisis As Integer
    ' [/GS]
    
    ' [GS]
    Magic As String * 16
    Poder As String * 16
    Agilidad As String * 16
    ' [/GS]
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    MinInt As Integer
    
    Real As Integer
    Caos As Integer
    
    ' v0.12a9
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    
    
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS = 40
'[/KEVIN]

'[KEVIN]
Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Type tReputacion 'Fama del usuario
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double
End Type

' [GS] Sistema administrativo
Type tAdmin
    Activado As Boolean
    EnPrueba As Boolean
    MaxCP As Integer
    Config As String
    CP(1 To 512) As String
End Type
' [/GS]

'Estadisticas de los usuarios
Type UserStats
    GLD As Long 'Dinero
    banco As Long
    MET As Integer
    
    MaxHP As Integer
    MinHP As Integer
    
    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    Def As Integer
    exp As Double
    ELV As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
End Type

'Flags
Type UserFlags
    ' [GS]
    Cliente As String
    ' [/GS]
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeAtacar As Byte
    PuedeMoverse As Byte
    QuiereLanzarSpell As Byte ' 0.12b1
    PuedeLanzarSpell As Byte
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    ' [NEW]
    Casado As String
    Casandose As String
    ' [/NEW]
    ' [GS]
    InvitaParty As Integer  ' Si esta invitando a alguien
'    Party As Integer       ' Numero de con quien este en party
    
    ' [GS]
    YaVoto As Boolean
    ' [/GS]
    
    ' [GS] Magias Explosivas
    NumHechExp As Integer   ' Numero del Hechi, explosivo
    TimerExp As Long        ' Timer de explociones
    TiraExp As Boolean      ' Si esta explotando ahora
    XExp As Integer         ' X del centro de la explocion
    YExp As Integer         ' Y del centro de la explocion
    
'    PartyInvito As Integer ' Ultimo invitado
    Partys(1 To 5) As Integer
    LiderParty As Integer
    
    BugLageador As Integer
    BorrarAlSalir As Boolean
    PocionRepelente As Boolean
    TiempoOnline As Long ' Minutos online
    ' [/GS]
    
    ' [GS] Sistema anti-logeo
    TiempoIni As Long
    RecienIni As Boolean
    ' [/GS]
    
    ' [GS] Counter
    CS_Esta As Boolean
    ' [/GS]
    
    ' [GS] Aventurero pr0
    AV_Esta As Boolean
    AV_Lugar As String
    AV_Tiempo As Integer
    ' [/GS]
    
    ' [GS] Reduce lag
    UltimoEST As String
    ' [/GS]
    
    ' [GS] Ultimo Nick & Color (Reparacion de salir/entrar)
    UltimoNickColor As String
    ' [/GS]
    
    ' [GS] Tiene mensaje arriba?
    TieneMensaje As Boolean
    ' [/GS]
    
    ' [GS] NPC q ataca
    SuNPC As Integer
    ' [/GS]
    
    Vuela As Byte
    Navegando As Byte
    Seguro As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As Integer ' Tipo del npc señalado
    NpcInv As Integer
    
    ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    
    StatsChanged As Byte
    Privilegios As Byte
    
    ' [GS] Ayudantes
    Ayudante As Boolean
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    
    '[Barrin 30-11-03] Hiper-AO
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    Trabajando As Boolean
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04] Hiper-AO
    UltimoMensaje As Byte
    '[/CDT]
    
    ' v0.12a9
    Mimetizado As Byte
    
    ' v0.12b1
    EsRolesMaster As Boolean
    PertAlCons As Boolean
    PertAlConsCaos As Boolean
    
    ' v0.12b2
    UsandoCodecXXX As Boolean
    
End Type

Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    Invisibilidad As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    '[Sicarul]
    AntiSH As Integer
    AntiSH2 As Integer
    '[/Sicarul]
    
    ' v0.12a9
    Mimetismo As Integer
End Type

Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    ' 0.12b1
    Reenlistadas As Byte
End Type

Type tGuild
    GuildName As String
    Solicitudes As Long
    SolicitudesRechazadas As Long
    Echadas As Long
    VecesFueGuildLeader As Long
    YaVoto As Byte
    EsGuildLeader As Byte
    FundoClan As Byte
    ClanFundado As String
    ClanesParticipo As Long
    GuildPoints As Double
    BorroClan As Boolean
End Type

'Tipo de los Usuarios
Type User
    
    Name As String
    ID As Long
    
    modName As String
    Password As String
    'wag Hiper-AO
        CheatCont As Integer
        Epa As Integer
    '/wag
    Char As Char 'Define la apariencia
    OrigChar As Char
    

    Administracion As tAdmin
    
    desc As String ' Descripcion
    clase As Byte
    raza As Byte
    genero As Byte
    Email As String
    Hogar As String
    
    ' [GS] NUEVO: Idioma?
    ' Idioma As String
    ' Si pone ENG, es en ingles
    ' lo demas es Español por ahora
    ' [/GS]
    
    Invent As Inventario
    
    Pos As WorldPos
    
    
    ConnID As Integer 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    NumeroPaquetesPorMiliSec As Long
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    
    Reputacion As tReputacion
    
    Faccion As tFacciones
    GuildInfo As tGuild
    GuildRef  As cGuild
    
    PrevCRC As Long
    PacketNumber As Long
    RandKey As Long
    
    IP As String
    
     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
    
    AntiCuelgue As Long
    
    'v0.12a9
    CharMimetizado As Char
    
    'v0.12b1
    Silenciado As Boolean
    NoExiste As Boolean
End Type




'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    Def As Integer
    UsuariosMatados As Integer
    ImpactRate As Integer
    LastEntrenar As Double ' Hiper-AO no lo tiene
End Type

Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
    
End Type

Type NPCFlags
    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    ' [GS] Barrera Espejo?
    BarreraEspejo As Byte
    ' [/GS]
    ' [GS] AtacaInvis
    AtacaInvis As Boolean
    ' [/GS]
    
    OldMovement As Byte
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    AttackedIndex As Integer
    Category1 As String
    Category2 As String
    Category3 As String
    Category4 As String
    Category5 As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    ' [GS] Hijos
    Hijo1 As Integer
    Hijo2 As Integer
    Hijo3 As Integer
    ' [/GS]
    
    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Snd4 As Integer
    
    ' [GS] Hablo?
    Hablo As Boolean
    Dijo As Byte ' NPC de Agite Only
    ' [/GS]
    
    'v0.12a9
    Inmovilizado As Byte
    
End Type

Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

'<--------- New type for holding the pathfinding info ------>
Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
'<--------- New type for holding the pathfinding info ------>


Type Npc
    Name As String
    Char As Char 'Define como se vera
    Equip As tEquip
    desc As String
    
    NPCtype As Integer
    Numero As Integer
    
    ' [GS] Tiene Mana???
    TieneMana As Boolean
    mana As Integer
    MiMana As Integer
    Meditando As Boolean
    ' [/GS]
    
    ' [GS] Tira equipamiento?
    TiraEquip As Boolean
    ' [/GS]
    
    ' [GS] NPC anti magias
    NoMagias As Byte
    ' [/GS]
    
    level As Integer
    
    InvReSpawn As Byte
    
    Comercia As Integer
    ' [GS] Intercambia
    Intercambia As Integer
    ' [/GS]
    ' [GS] Combate?
    Combate As Integer
    TempSum As Integer
    ' [/GS]
    
    Target As Long
    TargetNPC As Long
    TipoItems As Integer
    
    Veneno As Byte
    
    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer
    
    Movement As Integer
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    
    Inflacion As Long
    
    GiveEXP As Long
    GiveGLD As Long
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    '<---------New!! Needed for pathfindig----------->
    PFINFO As NpcPathFindingInfo

    
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    trigger As Integer
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean
    
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    
    ' [GS]
    Cargado As Boolean
    ' Modificado?
    Bloqueos As Boolean
    Triggers As Boolean
    Telep As Boolean
    NPCs As Boolean
    Objs As Boolean
    Datos As Boolean
    ' [/GS]
    
    ' v0.12b3
    MagiaSinEfecto As Byte
    
End Type



'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public ULTIMAVERSION2 As String
Public BackUp As Boolean

Public ListaRazas() As String
Public SkillsNames() As String
Public ListaClases() As String


Public ENDL As String
Public ENDC As String

Public RecordUsuarios As Long

'Directorios
Public IniPath As String
Public CharPath As String
Public MapPath As String
Public DatPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos
Public StartPos As WorldPos 'Posicion de comienzo


Public NumUsers As Integer 'Numero de usuarios actual
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public Oscuridad As Integer
Public NocheDia As Integer
Public PuedeCrearPersonajes As Byte

'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist() As Npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public MD5s() As String
Public BanIps As New Collection
'*********************************************************
' [GS]
Public Microfono As Integer
' [/GS]

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos


Public Ayuda As New cCola

Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GenCrC Lib "crc" Alias "GenCrc" (ByVal CrcKey As Long, ByVal CrcString As String) As Long


Sub PlayWaveAPI(file As String)

On Error Resume Next
Dim rc As Integer

rc = sndPlaySound(file, SND_ASYNC)

End Sub

