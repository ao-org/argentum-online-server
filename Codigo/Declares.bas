Attribute VB_Name = "Declaraciones"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
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

''
' Modulo de declaraciones. Aca hay de todo.
'

Public Enum e_elecciones
    HayGanador = 1
    HayGanadorPeroAbandono = 2
    HuboEmpate = 3
    NoVotos = 4
    AbroElecciones = 5
End Enum


Public Enum e_Facciones
    Criminal = 0
    Ciudadano = 1
    Caos = 2
    Armada = 3
End Enum

Public lstUsuariosDonadores() As String

Public Administradores As clsIniManager

Public Enum e_SoundIndex

    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 74
    FLECHA_IMPACTO = 65
    CONVERSION_BARCO = 55
    SOUND_COMIDA = 7

End Enum

Public Md5Cliente           As String

Public HoraMundo            As Long

Public HoraActual           As Integer

Public UltimoChar           As String

Public ExpMult              As Integer

Public OroMult              As Integer

Public DropMult             As Integer

Public RecoleccionMult      As Integer

Public DiceMinimum          As Integer

Public DiceMaximum          As Integer

Public EventoExpMult        As Integer

Public EventoOroMult        As Integer

Public EstadoGlobal         As Boolean

Public TimerLimpiarObjetos  As Byte

Public DuracionDia          As Long

Public OroPorNivel          As Integer

Public DropActive           As Byte

Public CuentaRegresivaTimer As Byte

Public PENDIENTE            As Integer

Public CostoPerdonPorCiudadano As Long

Public MaximoSpeedHack      As Integer

Public LastCountUsersOnline As Integer
Public LastRecordUsuarios   As Integer

Type t_EstadisticasDiarias

    segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer

End Type
    
Public DayStats       As t_EstadisticasDiarias

Public aClon          As New clsAntiMassClon

Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 15

Public Const INFINITE_LOOPS As Integer = -1

Public Const FXSANGRE = 14

''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL   As Long = &HF82FF

''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND          As Byte = 0

Public Const iFragataFantasmal = 87 'ok

Public Const iTraje = 694 'Traje +25
Public Const iTrajeAltoNw = 1295 'Traje -25 alto
Public Const iTrajeBajoNw = 1296 'Traje -25 enano

Public Const iBarca = 84

Public Const iBarcaCiuda = 1265

Public Const iBarcaCrimi = 1266

Public Const iGalera = 85

Public Const iGaleraCiuda = 1267

Public Const iGaleraCrimi = 1268

Public Const iGaleon = 86

Public Const iGaleonCiuda = 1269

Public Const iGaleonCrimi = 1270

Public Const iBarcaArmada = 1273

Public Const iBarcaCaos = 1274

Public Const iGaleraArmada = 1271

Public Const iGaleraCaos = 1272

Public Const iGaleonArmada = 1264

Public Const iGaleonCaos = 1263

Public Const iRopaBuceoMuerto = 772

Public MapasInterdimensionales() As Integer

Public MapasIgnoranLimpieza() As Integer

Public Enum e_Minerales

    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388

End Enum

Public Type t_LlamadaGM

    Usuario As String * 255
    Desc As String * 255

End Type

Public Enum e_PlayerType

    user = &H1
    RoleMaster = &H2
    Consejero = &H4
    SemiDios = &H8
    Dios = &H10
    Admin = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80

End Enum

Public Enum e_Class

    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Bard        'Bardo
    Druid       'Druida
    Paladin     'Paladín
    Hunter      'Cazador
    Trabajador  'Trabajador
    Pirat       'Pirata
    Thief       'Ladron
    Bandit      'Bandido
    
End Enum

Public Enum e_Ciudad

    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
    cArkhein

End Enum

Public Enum e_Raza

    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
    Orco

End Enum

Enum e_Genero

    Hombre = 1
    Mujer

End Enum

Public Enum e_ClanType

    ct_Neutral
    ct_ArmadaReal
    ct_LegionOscura
    ct_GM

End Enum

Public Const LimiteNewbie As Byte = 12

Public Type t_Cabecera 'Cabecera de los con

    Desc As String * 255
    crc As Long
    MagicWord As Long

End Type

Public MiCabecera                    As t_Cabecera

Public Const NingunEscudo            As Integer = 2

Public Const NingunCasco             As Integer = 2

Public Const NingunArma              As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 402

Public Const MAXMASCOTASENTRENADOR   As Byte = 7

Public Enum e_FXSound

    Lobo_Sound = 124
    Gallo_Sound = 137
    Dropeo_Sound = 132
    Casamiento_sound = 161
    BARCA_SOUND = 202
    MP_SOUND = 522

End Enum

Public Enum e_FXIDs
    FXWARP = 30
    FXMEDITARGRANDE = 42
End Enum

Public Enum e_Meditaciones
    MeditarInicial = 115
    MeditarMayor15 = 116
    MeditarMayor30 = 117
    MeditarMayor40 = 118
    MeditarMayor45 = 119
    MeditarMayor47 = 120
End Enum

Public Enum e_ParticulasIndex ' Particulas FX

    Envenena = 32
    Incinerar = 6
    Intermundia = 16
    Resucitar = 22
    Curar = 23
    LogeoLevel1 = 177
    CurarCrimi = 12
    Paralizar = 27
    Runa = 167
    TpVerde = 229

End Enum

Public Const VelocidadNormal       As Single = 1

Public Const VelocidadMuerto       As Single = 1.4

Public Const TIEMPO_CARCEL_PIQUETE As Long = 5

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum e_Trigger

    nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZonaSegura = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    CURA = 7
    DETALLEAGUA = 8
    CARCEL = 9
    PESCAVALIDA = 10
    
End Enum

''
' constantes para el trigger 6
'
' @see e_Trigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum e_Trigger6

    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3

End Enum

'TODO : Reemplazar por un enum
Public Const Bosque = "BOSQUE"

Public Const Nieve = "NIEVE"

Public Const Desierto = "DESIERTO"

Public Const Ciudad = "CIUDAD"

Public Const Campo = "CAMPO"

Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum e_TargetType

    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4

End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum e_TipoHechizo

    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3
    uInvocacion = 4
    uArea = 5
    uPortal = 6
    UFamiliar = 7
    uCombinados = 8
    
End Enum

Public Const MAX_MENSAJES_FORO As Byte = 35

Public Const MAXUSERHECHIZOS   As Byte = 25

Public Const FX_TELEPORT_INDEX             As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum e_PartesCuerpo

    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6

End Enum

Public Const Guardias As Integer = 6

Public Const MAX_PERSONAJES = 10

Public Const MAXORO         As Long = 90000000

Public Const MAXEXP         As Long = 1999999999

Public Const MAXUSERMATADOS As Long = 65000

Public Const MINATRIBUTOS   As Byte = 6

Public Const LingoteHierro  As Integer = 386 'OK

Public Const LingotePlata   As Integer = 387 'OK

Public Const LingoteOro     As Integer = 388 'OK

Public Const Leña           As Integer = 58 'OK

Public Const LeñaElfica     As Integer = 2781 'OK

Public Const Raices         As Integer = 888 'OK

Public Const PieldeLobo     As Integer = 414 'OK

Public Const PieldeOsoPardo As Integer = 415 'OK

Public Const PieldeOsoPolar As Integer = 416 'OK

Public Const MaxNPCs        As Integer = 10000

Public Const MAXCHARS       As Integer = 10000

Public Const DAGA                As Integer = 15 'OK

Public Const FOGATA_APAG         As Integer = 136 'OK

Public Const FOGATA              As Integer = 63 'OK

Public Const ORO_MINA            As Integer = 194 'OK

Public Const PLATA_MINA          As Integer = 193 'OK

Public Const HIERRO_MINA         As Integer = 192 'OK

Public Const ObjArboles          As Integer = 4 'OK


Public Enum e_NPCType

    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Enlistador = 5
    DRAGON = 6
    Timbero = 7
    GuardiasCaos = 8
    ResucitadorNewbie = 9
    Pirata = 10
    Veterinaria = 11
    Gobernador = 12
    GuardiaNpc = 13
    Subastador = 16
    Quest = 17
    Pretoriano = 18
    DummyTarget = 19
    
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS      As Byte = 24

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS   As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES      As Byte = 12

''
' Cantidad de Razas
Public Const NUMRAZAS       As Byte = 5

''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS    As Byte = 3

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum e_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Enum e_Block

    NORTH = &H1
    EAST = &H2
    SOUTH = &H4
    WEST = &H8
    ALL_SIDES = &HF
    GM = &H10

End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 829

Public Const iCabezaMuerto As Integer = 0 ' El nuevo casper no usa cabeza. El viejo es: 621

Public Const iORO          As Byte = 12

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum e_Skill

    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Comerciar = 9
    Defensa = 10
    liderazgo = 11
    Proyectiles = 12
    Wrestling = 13
    Navegacion = 14
    equitacion = 15
    Resistencia = 16

    Talar = 17
    Pescar = 18
    Mineria = 19
    Herreria = 20
    Carpinteria = 21
    Alquimia = 22
    Sastreria = 23
    Domar = 24

    Grupo = 90
    MarcaDeClan = 91
    MarcaDeGM = 92

End Enum

Public Const FundirMetal = 88

Public Enum e_Atributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Constitucion = 4
    Carisma = 5

End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel

Public Const AdicionalHPCazador  As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef        As Byte = 15

Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3

Public Const AumentoSTMago       As Byte = AumentoSTDef - 1

Public Const AumentoStBandido    As Byte = AumentoSTDef + 3

'Tamaño del mapa
Public Const XMaxMapSize         As Byte = 100

Public Const XMinMapSize         As Byte = 1

Public Const YMaxMapSize         As Byte = 100

Public Const YMinMapSize         As Byte = 1

'Tamaño del tileset
Public Const TileSizeX           As Byte = 32

Public Const TileSizeY           As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow             As Byte = 23

Public Const YWindow             As Byte = 18

'Sonidos
Public Const SND_SWING           As Byte = 2

Public Const SND_TALAR           As Byte = 13

Public Const SND_TIJERAS         As Byte = 211

Public Const SND_PESCAR          As Byte = 14

Public Const SND_MINERO          As Byte = 15

Public Const SND_WARP            As Byte = 3

Public Const SND_PUERTA          As Integer = 5

Public Const SND_PUERTA_DUCTO    As Integer = 380

Public Const SND_NIVEL           As Integer = 554

Public Const SND_USERMUERTE      As Byte = 11

Public Const SND_IMPACTO         As Byte = 10

Public Const SND_IMPACTO_APU     As Integer = 2187

Public Const SND_IMPACTO_CRITICO As Integer = 2186

Public Const SND_IMPACTO2        As Byte = 12

Public Const SND_DOPA As Byte = 77

Public Const SND_LEÑADOR         As Byte = 13

Public Const SND_FOGATA          As Byte = 116

Public Const SND_SACARARMA       As Byte = 25

Public Const SND_ESCUDO          As Byte = 37

Public Const MARTILLOHERRERO     As Byte = 41

Public Const LABUROCARPINTERO    As Byte = 42

Public Const SND_BEBER           As Byte = 135

Public Const GRH_FALLO_PESCA     As Long = 48974

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS      As Integer = 10000

''
' Cantidad de "slots" en el inventario con todos los slots desbloqueados
Public Const MAX_INVENTORY_SLOTS     As Byte = 42

' Cantidad de "slots" en el inventario básico
Public Const MAX_USERINVENTORY_SLOTS As Byte = 24

' Cantidad de "slots" en el inventario por fila
Public Const SLOTS_PER_ROW_INVENTORY As Byte = 6

' Cantidad máxima de filas a desbloquear en el inventario
Public Const INVENTORY_EXTRA_ROWS    As Byte = 3

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO                 As Integer = 200

Public Const FLAG_AGUA               As Byte = &H20

Public Const FLAG_ARBOL              As Byte = &H40

' CATEGORIAS PRINCIPALES
Public Enum e_OBJType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otEscudo = 16
    otCasco = 17
    otHerramientas = 18
    otTeleport = 19
    OtDecoraciones = 20
    otMagicos = 21
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otDañoMagico = 30
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otResistencia = 35
    otpasajes = 36
    otmapa = 38
    OtPozos = 40
    otMonturas = 44
    otRunas = 45
    otNudillos = 46
    OtCorreo = 47
    OtCofre = 48
    OtDonador = 50
    otCualquiera = 1000

End Enum

'Estadisticas
Public Const STAT_MAXELV              As Byte = 47

Public Const STAT_MAXHP               As Integer = 32000

Public Const STAT_MAXMP               As Integer = 32000

Public Const STAT_MAXSTA              As Integer = 32000

Public Const STAT_MAXMAN              As Integer = 32000

Public Const STAT_MAXHIT              As Integer = 999

Public Const STAT_MAXDEF              As Byte = 99

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type t_Hechizo
    AutoLanzar As Byte
    
    velocidad As Single
    Duration As Integer
    RequiredHP As Integer
    
    CoolDown As Integer
    
    ScreenColor As Long
    TimeEfect As Long

    'Hechizo de teleport
    TeleportX As Integer
    TeleportXMap As Integer
    TeleportXX As Integer
    TeleportXY As Integer

    'Hechizo de Materialización
    MaterializaObj As Integer
    MaterializaCant As Integer
    
    NecesitaObj As Integer
    NecesitaObj2 As Integer
    'Hechizos de area
    
    AreaRadio As Long
    AreaAfecta As Integer
    
    Particle As Integer
    TimeParticula As Integer
    ParticleViaje As Integer
    desencantar As Byte
    Sanacion As Byte
    AntiRm As Byte
    
    'Sistema..
    nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String

    Tipo As e_TipoHechizo
    wav As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    
    SubeMana As Byte
    MinMana As Integer
    MaxMana As Integer
    
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
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    incinera As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    GolpeCertero As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    RemueveInvisibilidadParcial As Byte
    
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

    '    Materializa As Byte
    '    ItemIndex As Byte
    
    Mimetiza As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As e_TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean

End Type

Public Type t_UserOBJ

    ObjIndex As Integer
    amount As Integer
    Equipped As Byte

End Type

Public Type t_Inventario

    Object(1 To MAX_INVENTORY_SLOTS) As t_UserOBJ
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
    DañoMagicoEqpObjIndex As Integer
    DañoMagicoEqpSlot As Byte
    ResistenciaEqpObjIndex As Integer
    ResistenciaEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    MagicoObjIndex As Integer
    MagicoSlot As Byte
    NudilloObjIndex As Integer
    NudilloSlot As Byte
    
End Type

Public Type t_WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type t_Position

    X As Integer
    Y As Integer

End Type

Public Type t_CityWorldPos

    Map As Integer
    X As Integer
    Y As Integer
    MapaViaje As Integer
    ViajeX As Byte
    ViajeY As Byte
    MapaResu As Integer
    ResuX As Byte
    ResuY As Byte
    NecesitaNave As Byte

End Type

Public Type t_FXdata

    nombre As String
    GrhIndex As Long
    Delay As Integer

End Type

'Datos de user o npc
Public Type t_Char

    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    ParticulaFx As Integer
    FX As Integer
    loops As Integer
    Heading As e_Heading
    Head_Aura As String
    Body_Aura As String
    Arma_Aura As String
    Escudo_Aura As String
    DM_Aura As String
    RM_Aura As String
    Otra_Aura As String
    speeding As Single
    BodyIdle As Integer
    Animation() As Integer

End Type

Public Type t_Obj

    ObjIndex As Integer
    amount As Long
    Data As Double

End Type

Public Type t_QuestNpc

    NpcIndex As Integer
    amount As Integer

End Type

Public Type t_QuestObj

    QuestIndex As Integer
    ObjIndex As Integer
    amount As Integer
    Probabilidad As Long

End Type
 
Public Type t_UserQuest

    NPCsTarget() As Integer
    NPCsKilled() As Integer
    QuestIndex As Integer

End Type

Public QuestList() As t_Quest

Public Type t_Quest

    nombre As String
    Desc As String
    NextQuest As String
    DescFinal As String
    RequiredLevel As Byte
    
    RequiredQuest As Byte
    
    RequiredOBJs As Byte
    RequiredOBJ() As t_Obj
    
    RequiredNPCs As Byte
    RequiredNPC() As t_QuestNpc
    
    
    RequiredTargetNPCs As Byte
    RequiredTargetNPC() As t_QuestNpc
    
    RewardGLD As Long
    RewardEXP As Long
    
    RewardOBJs As Byte
    RewardOBJ() As t_Obj
    Repetible As Byte

End Type

' ******************* RETOS ************************
Public Enum e_SolicitudRetoEstado
    Libre
    Enviada
    EnCola
End Enum

Public Type t_SolicitudJugador
    nombre As String
    Aceptado As Boolean
    CurIndex As Integer
End Type

Public Type t_SolicitudReto
    estado As e_SolicitudRetoEstado
    Jugadores() As t_SolicitudJugador
    Apuesta As Long
    PocionesMaximas As Integer
    CaenItems As Boolean
End Type

Public Enum e_EquipoReto
    Izquierda
    Derecha
End Enum

Public Type t_SalaReto
    PosIzquierda As t_WorldPos
    PosDerecha As t_WorldPos
    IndexBanquero As Integer
    ' -----------------
    EnUso As Boolean
    Ronda As Byte
    Puntaje As Integer
    Apuesta As Long
    PocionesMaximas As Integer
    CaenItems As Boolean
    TiempoRestante As Long
    TiempoItems As Integer
    TamañoEquipoIzq As Byte
    TamañoEquipoDer As Byte
    Jugadores() As Integer
End Type

Public Type t_Retos
    TamañoMaximoEquipo As Byte
    ApuestaMinima As Long
    ImpuestoApuesta As Single
    DuracionMaxima As Long
    TiempoConteo As Byte
    Salas() As t_SalaReto
    TotalSalas As Integer
    SalasLibres As Integer
    AnchoSala As Integer
    AltoSala As Integer
End Type
' **************************************************

'Tipos de objetos
Public Type t_ObjData
    Elfico As Byte
    velocidad As Single
    CantEntrega As Byte
    CantItem As Byte
    Item() As t_Obj
    ParticulaGolpeTime As Integer
    ParticulaGolpe As Integer
    ParticulaViaje As Integer

    ClaseTipo As Byte

    TipoRuna As Byte

    Name As String 'Nombre del obj
    
    OBJType As e_OBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As t_Inventario
    Apuñala As Byte
    
    Paraliza As Byte
    Estupidiza As Byte
    Envenena As Byte
    
    NoSeLimpia As Byte
    Subastable As Integer
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHp As Integer ' Minimo puntos de vida
    MaxHp As Integer ' Maximo puntos de vida
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    Proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    
    ' Sistema de armas Dos Manos - SimP - 03/02/2021
    DosManos As Byte
    
    Newbie As Integer
    
    'By Ladder
    CreaParticula As String
    
    CreaFX As Integer
    
    CreaLuz As String
    
    MinELV As Byte
    SkillIndex As Byte     ' El indice de Skill para equipar el item
    SkillRequerido As Byte ' El valor MINIMO requerido de skillIndex para equipar el item
    
    CreaGRH As String
    SndAura As Integer
    Intirable As Byte
    Instransferible As Byte
    Destruye As Byte
    
    NecesitaNave As Byte

    DesdeMap As Long
    HastaMap As Long
    HastaY As Byte
    HastaX As Byte
    
    EfectoMagico As Byte
    Que_Skill As Byte          ' Que skill recibe la bonificacion
    CantidadSkill As Byte     ' Cuantos puntos de skill bonifica
    
    Subtipo As Byte ' 0: -, 1: Paraliza, 2: Incinera, 3: Envenena, 4: Explosiva
    
    Dorada As Byte
    
    VidaUtil As Integer
    TiempoRegenerar As Integer
    
    CuantoAumento As Single ' Cuanto aumenta el atributo.
    QueAtributo As Byte     ' Que attributo sube (Agilidad, Fuerza, etc)
    incinera As Byte

    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    Porcentaje As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHit As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    RazaOrca As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
    Raices As Integer
    SkPociones As Byte
    
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolaR As Integer
    SkMAGOria As Byte
     
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As e_Class
    
    'Razas que no tienen permitido usar este obj
    RazaProhibida(1 To NUMRAZAS) As e_Raza
    
    ClasePermitida As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    
    Power As Integer
    MagicDamageBonus As Integer
    ResistenciaMagica As Integer
    Revive As Boolean
    Refuerzo As Byte

    Invernal As Boolean
    
    CatalizadorTipo As Byte
    CatalizadorAumento As Single

    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
End Type

'[Pablo ToxicWaste]
' Mod. by WyroX
Public Type t_ModClase

    Vida As Double
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    ModApuñalar As Double
    Escudo As Double
    ManaInicial As Double
    AumentoSta As Integer
    MultMana As Double
    HitPre36 As Integer
    HitPost36 As Integer
    ResistenciaMagica As Double

End Type

Public Type t_ModRaza

    Fuerza As Integer
    Agilidad As Integer
    Inteligencia As Integer
    Carisma As Integer
    Constitucion As Integer

End Type

'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 42

'[/KEVIN]

'[KEVIN]
Public Type t_BancoInventario

    Object(1 To MAX_BANCOINVENTORY_SLOTS) As t_UserOBJ
    NroItems As Integer

End Type

Public Enum e_TipoUsuario
    tNormal = 0
    tCafecito
    tHeroe
    tAventurero
    tLeyenda
End Enum

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

'Estadisticas de los usuarios
Public Type t_UserStats

    tipoUsuario As e_TipoUsuario
    GLD As Long 'Dinero
    InventLevel As Byte 'Filas extra desbloqueadas en el inventario
    Banco As Long
    
    MaxHp As Integer
    MinHp As Integer
    
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHit As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Long
    ELV As Byte
    ELO As Long

    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    CriminalesMatados As Long
    NPCsMuertos As Integer
    SkillPts As Integer
    Advertencias As Byte
End Type

'Sistema de Barras
Public Type t_AccionPendiente

    AccionPendiente As Boolean
    TipoAccion As e_AccionBarra
    RunaObj As Integer
    ObjSlot As Byte
    Particula As Byte
    HechizoPendiente As Integer

End Type

'Flags
Public Type t_UserFlags
    Nadando As Byte
    
    NecesitaOxigeno As Boolean

    Ahogandose As Byte
    EnTorneo As Boolean

    'Ladder
    'Casamientos  08/6/10 01:10 am
    Pareja As String
    Casado As Byte
    Candidato As Integer
    
    pregunta As Byte
       
    ' 0: no esta hechizada;
    'Cualquier otro valor si lo esta: 0.8 -> reduce un 20% de velocidad; 1.3 -> Aumenta un 30%
    VelocidadHechizada As Single
    
    LevelBackup As Byte
    
    UsandoMacro As Boolean
    PendienteDelSacrificio As Byte
    PendienteDelExperto As Byte
    NoPalabrasMagicas As Byte
    incinera As Byte
    Envenena As Byte
    Paraliza As Byte
    Estupidiza As Byte
    NoMagiaEfecto As Byte
    CarroMineria As Byte
    GolpeCertero As Byte
    AnilloOcultismo As Byte
    NoDetectable As Byte
    RegeneracionMana As Byte
    RegeneracionHP As Byte
    RegeneracionSta As Byte
    
    'Hechizo de Transportacion
    
    Portal As Integer
    PortalM As Integer
    PortalX As Integer
    PortalY As Integer
    PortalMDestino As Integer
    PortalXDestino As Integer
    PortalYDestino As Integer
    'Hechizo de Transportacion
    
    
    Inmunidad As Byte
    
    Inmovilizado As Byte
    
    Montado As Byte
    Subastando As Boolean
    Incinerado As Byte
    'Ladder
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    FirstPacket As Boolean ' ¿El socket envió algun paquete válido?
    Meditando As Boolean
    Escribiendo As Boolean
    Crafteando As Byte

    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    Mimetizado As e_EstadoMimetismo
    MascotasGuardadas As Byte
    
    Navegando As Byte
    
    Seguro As Boolean
    SeguroParty As Boolean
    SeguroClan As Boolean
    SeguroResu As Boolean

    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As e_NPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    BanMotivo As String

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
    NPCAtacado As Integer
    
    StatsChanged As Byte
    Privilegios As e_PlayerType
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    VecesQueMoriste As Long
    MinutosRestantes As Byte
    SegundosPasados As Byte
    
    ChatColor As Long

    UltimoMensaje As Integer
    
    Silenciado As Byte
    
    'Centinela
    CentinelaOK As Boolean
    
    Traveling As Byte
    
    EnConsulta As Boolean
    
    ProcesosPara As String
    ScreenShotPara As String

    ChatHistory(1 To 15) As String
    
    EnReto As Boolean
    SalaReto As Integer
    EquipoReto As e_EquipoReto
    AceptoReto As Integer
    SolicitudReto As t_SolicitudReto
    LastPos As t_WorldPos
    
    ReturnPos As t_WorldPos

    YaGuardo As Boolean
    
    ModificoAttributos As Boolean
    ModificoHechizos As Boolean
    ModificoInventario As Boolean
    ModificoInventarioBanco As Boolean
    ModificoSkills As Boolean
    ModificoMascotas As Boolean
    ModificoQuests As Boolean
    ModificoQuestsHechas As Boolean

End Type

Public Enum e_EstadoMimetismo
    Desactivado = 0
    FormaUsuario = 1
    FormaBichoSinProteccion = 2
    FormaBicho = 3
End Enum

Public Type t_UserCounters

    TiempoDeInmunidad As Byte
    EnCombate As Byte

    TiempoParaSubastar As Byte
    UserHechizosInterval(1 To MAXUSERHECHIZOS) As Long
    Oxigeno As Long
    
    Ahogo As Long
    
    IdleCount As Integer
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Incineracion As Integer
    Paralisis As Integer
    velocidad As Integer
    Inmovilizado As Integer
    Ceguera As Integer
    Estupidez As Integer
    Mimetismo As Integer

    ' Anticheat
    SpeedHackCounter As Single
    LastStep As Long
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    
    PiqueteC As Long
    Pena As Long
    SendMapCounter As t_WorldPos
    Pasos As Integer
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    RepetirMensaje As Integer
    MensajeGlobal As Long
    
    Maldicion As Byte

    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerCaminar As Long
    TimerTirar As Long
    TimerMeditar As Long

    'Nuevos de AoLibre
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long

    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela

    goHome As Long

    LastSave As Long
    
    CuentaRegresiva As Integer
    TimerBarra As Integer

End Type

Public Type t_UserIntervals

    Magia As Long
    Golpe As Long
    Arco As Long
    UsarU As Long
    UsarClic As Long
    Caminar As Long
    GolpeMagia As Long
    MagiaGolpe As Long
    GolpeUsar As Long
    TrabajarExtraer As Long
    TrabajarConstruir As Long

End Type

Public Type t_QuestStats

    Quests(1 To MAXUSERQUESTS) As t_UserQuest
    NumQuestsDone As Integer
    QuestsDone() As Integer

End Type

' ------------- FACCIONES -------------

Public Type t_Facciones

    Status As Byte ' Esto deberia ser e_Facciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Long
    ciudadanosMatados As Long
    RecompensasReal As Long ' a.k.a Rango armada real
    RecompensasCaos As Long ' a.k.a Rango legion caos
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    MatadosIngreso As Integer 'Para Armadas nada mas
    NextRecompensa As Integer 'DEPRECATED: Atributo viejo. Deberiamos usar `tRangoFaccion`

End Type

Public Type t_RangoFaccion

    rank As Byte
    Titulo As String
    NivelRequerido As Byte
    AsesinatosRequeridos As Integer

End Type

Public Type t_RecompensaFaccion

    rank As Byte
    ObjIndex As Integer

End Type


Public Type t_ComercioUsuario
    
    itemsAenviar(1 To 6) As t_Obj ' Mas de 6 no se puede, la UI muestra solo eso.
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    Oro As Long
    cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    Acepto As Boolean

End Type

Public Type t_UserTrabajo
    TargetSkill As e_Skill
    Target_X As Integer
    Target_Y As Integer
End Type

'Tipo de los Usuarios
Public Type t_User

    Name As String
    Cuenta As String
    
    ID As Long
    Trabajo As t_UserTrabajo
    AccountID As Long
    Grupo As Tgrupo

    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    Char As t_Char 'Define la apariencia
    CharMimetizado As t_Char
    NameMimetizado As String
    OrigChar As t_Char
    
    Desc As String ' Descripcion
    DescRM As String
    
    clase As e_Class
    raza As e_Raza
    genero As e_Genero
    Email As String
    Hogar As e_Ciudad
    PosibleHogar As e_Ciudad
    MENSAJEINFORMACION As String
        
    Invent As t_Inventario
    
    Pos As t_WorldPos
    
    ConnIDValida As Boolean
    
    CurrentInventorySlots As Byte
    
    BancoInvent As t_BancoInventario

    Counters As t_UserCounters
    Intervals As t_UserIntervals
    
    Stats As t_UserStats
    flags As t_UserFlags
    Accion As t_AccionPendiente

    Faccion As t_Facciones
    Familiar As t_Family

    ChatCombate As Byte
    ChatGlobal As Byte
    'Macros

    #If ConUpTime Then
        LogOnTime As Date
        UpTime As Long
    #End If

    IP As String
    
    '[Alejo]
    ComUsu As t_ComercioUsuario
    '[/Alejo]
    
    EmpoCont As Byte
    
    NroMascotas As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As e_ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    LastGuildRejection As String
    
    KeyCrypt As Integer
    
    AreasInfo As t_AreaInfo

    QuestStats As t_QuestStats

    Keys(1 To MAXKEYS) As Integer
    
    #If AntiExternos = 1 Then
        XorIndexIn As Integer
        XorIndexOut As Integer
    #End If

    CraftInventory(1 To MAX_SLOTS_CRAFTEO) As Integer
    CraftCatalyst As t_Obj
    CraftResult As clsCrafteo

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

Public Type t_NPCStats

    MaxHp As Long
    MinHp As Long
    MaxHit As Integer
    MinHIT As Integer
    def As Integer
    defM As Integer
    UsuariosMatados As Integer

End Type

Public Type t_NpcCounters

    Paralisis              As Long
    Inmovilizado           As Long
    TiempoExistencia       As Long
    IntervaloAtaque        As Long
    IntervaloMovimiento    As Long
    IntervaloLanzarHechizo As Long
    IntervaloRespawn       As Long
    UltimoAtaque           As Long

End Type

Public Type t_NPCFlags

    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As e_Facciones
    LanzaSpells As Byte
    NPCIdle As Boolean

    ' Invasiones
    InvasionIndex As Integer
    SpawnBox As Integer
    IndexInInvasion As Integer

    ExpCount As Long '[ALEJO]
    
    OldMovement As e_TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    ' UseAINow As Boolean No se usa, borrar de la DB!!!!
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    AttackedFirstBy As String
    backup As Byte
    RespawnOrigPos As Byte

    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Incinerado As Byte
    invisible As Byte
    Bendicion As Byte

    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer

    AtacaUsuarios As Boolean ' Si el NPC puede atacar usuarios
    AtacaNPCs As Boolean     ' Si el NPC puede atacar otros NPC

    AIAlineacion As e_Alineacion

End Type

Public Type t_CriaturasEntrenador

    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
    PuedeInvocar As Boolean

End Type

Public Type t_NpcPathFindingInfo

    PathLength As Integer   ' Number of steps *
    Path() As t_Vertice      ' This array holds the path
    destination As t_Position ' The location where the NPC has to go
    RangoVision As Single
    Inteligencia As Integer
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type

Public Type t_Caminata
    Offset As t_Position
    Espera As Long
End Type

Public Enum e_TipoAI
    Estatico = 1
    MueveAlAzar = 2
    NpcDefensa = 4
    SigueAmo = 8
    NpcAtacaNpc = 9
    GuardiaPersigueNpc = 10

    'Pretorianos
    SacerdotePretorianoAi = 11
    GuerreroPretorianoAi = 12
    MagoPretorianoAi = 13
    CazadorPretorianoAi = 14
    ReyPretoriano = 15

    ' Animado
    Caminata = 20
    
    ' Eventos
    Invasion = 21
End Enum

Public Enum e_Alineacion
    ninguna = 0
    Real = 1
    Caos = 2
End Enum

Public Type t_Npc
    
    Distancia As Byte
    
    NumDropQuest As Byte
    DropQuest() As t_QuestObj
    
    InformarRespawn As Byte
    Name As String
    SubName As String
    Char As t_Char 'Define como se vera
    Desc As String
    DescExtra As String
    showName As Byte
    GobernadorDe As Byte

    NPCtype As e_NPCType
    Numero As Integer

    nivel As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Craftea As Byte
    
    Target As Long
    TargetNPC As Long
    TipoItems As Integer
    
    SoundOpen As Integer
    SoundClose As Integer

    Veneno As Byte

    Pos As t_WorldPos 'Posicion
    Orig As t_WorldPos

    Movement As e_TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    
    GiveEXPClan As Long
    
    GiveEXP As Long
    GiveGLD As Long
    
    NumQuest As Integer
    QuestNumber() As Byte

    Stats As t_NPCStats
    flags As t_NPCFlags
    Contadores As t_NpcCounters
    IntervaloMovimiento As Long
    IntervaloAtaque As Long
    IntervaloLanzarHechizo As Long
    IntervaloRespawn As Long
    
    Invent As t_Inventario
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    ' Entrenadores
    NroCriaturas As Integer
    Criaturas() As t_CriaturasEntrenador
    EsFamiliar As Byte
    MaestroNPC As Integer
    MaestroUser As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    pathFindingInfo As t_NpcPathFindingInfo
    
    ' Esto es del Areas.bas
    AreasInfo As t_AreaInfo
    
    NumQuiza As Byte
    QuizaDropea() As String
    QuizaProb As Integer
    
    SubeSupervivencia As Byte
    
    NumDestinos As Byte
    Dest() As String
    Interface As Byte
    
    'Para diferenciar entre clanes
    ClanIndex As Integer
    
    Caminata() As t_Caminata
    CaminataActual As Byte
    
    PuedeInvocar As Byte
    
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type t_light

    Rango As Integer
    Color As Long

End Type

Public Type t_MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Long
    UserIndex As Integer
    NpcIndex As Integer
    Particula As Byte
    TimeParticula As Integer
    ObjInfo As t_Obj
    TileExit As t_WorldPos
    trigger As e_Trigger
    ParticulaIndex As Integer
    Luz As t_light

End Type

'Info del mapa
Type t_MapInfo

    map_name As String
    backup_mode As Byte
    music_numberHi As Long
    music_numberLow As Long
    Seguro As Byte
    zone As String
    terrain As String
    Newbie As Boolean
    SinMagia As Boolean
    SinInviOcul As Boolean
    NoPKs As Boolean
    NoCiudadanos As Boolean
    SoloClanes As Boolean
    ResuCiudad As Boolean
    ambient As String
    base_light As Long
    letter_grh As Long
    lluvia As Byte
    Nieve As Byte
    niebla As Byte
    NumUsers As Long
    MinLevel As Integer
    MaxLevel As Integer
    Salida As t_WorldPos
    NoMascotas As Boolean

End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE                      As Boolean

Public ULTIMAVERSION                     As String

Public backup                            As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS)         As String

Public SkillsNames(1 To NUMSKILLS)       As String

Public ListaClases(1 To NUMCLASES)       As String

Public ListaAtributos(1 To NUMATRIBUTOS) As String

Public RecordUsuarios                    As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath                           As String

Public CuentasPath                       As String

Public DeleteCuentasPath                 As String

''
'Ruta base para guardar los chars
Public CharPath                          As String

''
'Ruta base para guardar los users borrados
Public DeletePath                        As String

''
'Ruta base para los archivos de mapas
Public MapPath                           As String

''
'Ruta base para los DATs
Public DatPath                           As String

''
'Bordes del mapa
Public MinXBorder                        As Byte

Public MaxXBorder                        As Byte

Public MinYBorder                        As Byte

Public MaxYBorder                        As Byte

Public ResPos                            As t_WorldPos ' TODO: Se usa esta variable ?

''
'Numero de usuarios actual
Public NumCuentas                        As Long

Public NumUsers                          As Integer

Public LastUser                          As Integer

Public LastChar                          As Integer

Public NumChars                          As Integer

Public LastNPC                           As Integer

Public NumNPCs                           As Integer

Public NumMaps                           As Long

Public NumObjDatas                       As Integer

Public NumeroHechizos                    As Integer

Public MaxConexionesIP                   As Integer

Public MaxUsersPorCuenta                 As Byte

Public IdleLimit                         As Integer

Public MaxUsers                          As Integer

Public HideMe                            As Byte

Public MaxRangoFaccion                   As Byte ' El rango maximo que se puede alcanzar

Public LastBackup                        As String

Public minutos                           As String

Public haciendoBK                        As Boolean

Public PuedeCrearPersonajes              As Integer

Public ServerSoloGMs                     As Integer

Public EnPausa                           As Boolean

Public EnTesting                         As Boolean

'*****************ARRAYS PUBLICOS*************************
Public UserList()                         As t_User 'USUARIOS

Public NpcList(1 To MaxNPCs)              As t_Npc 'NPCS

Public MapData()                          As t_MapBlock

Public MapInfo()                          As t_MapInfo

Public Hechizos()                         As t_Hechizo

Public CharList(1 To MAXCHARS)            As Integer

Public ObjData()                          As t_ObjData

Public FX()                               As t_FXdata

Public SpawnList()                        As t_CriaturasEntrenador

Public ForbidenNames()                    As String

Public ArmasHerrero()                     As Integer

Public ArmadurasHerrero()                 As Integer

Public ObjCarpintero()                    As Integer

Public ObjAlquimista()                    As Integer

Public ObjSastre()                        As Integer

Public EspecialesTala()                   As t_Obj

Public EspecialesPesca()                  As t_Obj

Public Peces()                            As t_Obj

Public PesoPeces()                        As Long

Public RangosFaccion()                    As t_RangoFaccion

Public RecompensasFaccion()               As t_RecompensaFaccion

Public ModClase(1 To NUMCLASES)           As t_ModClase

Public ModRaza(1 To NUMRAZAS)             As t_ModRaza

Public Crafteos                           As New Dictionary
'*********************************************************

Public Nix                                As t_WorldPos

Public Ullathorpe                         As t_WorldPos

Public Banderbill                         As t_WorldPos

Public Lindos                             As t_WorldPos

Public Arghal                             As t_WorldPos

Public Arkhein                            As t_WorldPos

Public CityNix                            As t_CityWorldPos

Public CityUllathorpe                     As t_CityWorldPos

Public CityBanderbill                     As t_CityWorldPos

Public CityLindos                         As t_CityWorldPos

Public CityArghal                         As t_CityWorldPos

Public CityArkhein                        As t_CityWorldPos

Public Prision                            As t_WorldPos

Public Libertad                           As t_WorldPos

Public Ayuda                              As New cCola

Public ConsultaPopular                    As New ConsultasPopulares

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal Length As Long)

' Los Objetos Criticos nunca desaparecen del inventario de los npcs vendedores, una vez que
' se venden los 10.000 (max. cantidad de items x slot) vuelven a reabastecer.
Public Enum e_ObjetosCriticos

    Manzana = 1
    Manzana2 = 64
    ManzanaNewbie = 573

End Enum

#If AntiExternos = 1 Then
    Public Security As New clsSecurity
#End If

Public Type t_Rectangle
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
End Type

Public Const DIAMETRO_VISION_GUARDIAS_NPCS As Byte = 7

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum
