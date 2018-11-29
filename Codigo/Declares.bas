Attribute VB_Name = "Declaraciones"
'Lapsus AO 2009
'Copyright (C) 2009 Dalmasso, Juan Andres
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

''
' Modulo de declaraciones. Aca hay de todo.

Public UserSolicitadoFPS As String
Public MixedKey As Long
Public ServerIp As String
Public minPuntosClan As Long
Public CrcSubKey As String
Public CuentaRegresiva As Long
Type tEstadisticasDiarias
    segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer
End Type
    
Public DayStats As tEstadisticasDiarias

#If SeguridadAlkon Then
Public aDos As New clsAntiDoS
#End If

Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

'CHOTS | Fotos Remotas
Public Fotos_Ruta As String
Public Fotos_Longitud As Long
Public Fotos_Temporal As String

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14

'CHOTS, BysNacK | Seguridad AntiBoters
Public NumIps As Byte

'CHOTS | Sistema de Castillos

Public castNorte As String
Public castSur As String
Public castOeste As String
Public castEste As String

Public cAvisado As ccCastillos
Public ccAvisado As ccCastillos

Public Type ccCastillos
   Norte As Integer
   Sur As Integer
   Oeste As Integer
   Este As Integer
End Type

Public lastClan As lstClanCast

Public Type lstClanCast
   Norte As String
   Sur As String
   Oeste As String
   Este As String
End Type
   
'CHOTS | Sistema de Castillos

Public Const iFragataFantasmal = 87

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Type tLlamadaGM
    usuario As String * 255
    Desc As String * 255
End Type

Public Enum PlayerType
    User = 0
    Consejero = 1
    Ot = 2
    SemiDios = 3
    Dios = 4
End Enum

Public Const LimiteNewbie As Byte = 12

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'Barrin 3/10/03
Public Const TIEMPO_INICIOMEDITAR As Byte = 3

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 402
Public Const ExterminaIndex As Integer = 1005
Public Const BacuDragonIndex As Integer = 1006
Public Const LAUDMAGICO As Integer = 696
Public Const LAUDUNICO As Integer = 1133
Public Const FLAUTAMAGICA As Integer = 1082
Public Const FLAUTAUNICA As Integer = 1132
Public Const GEMAVIOLETA As Integer = 413
Public Const GEMACAPTURA As Integer = 1138 'CHOTS | Monturas
Public Const ROPAMONTURA As Integer = 1145 'CHOTS | Monturas

Public Const MAXMASCOTASENTRENADOR As Byte = 7

'CHOTS | Constantes de Meditaciones
Public Enum FXIDs
    FXWARP = 1
    FXCAPTURA = 47 'CHOTS | Monturas
    FXROTURA = 1 'CHOTS | Monturas
    FXMEDITAR1 = 16
    FXMEDITAR7 = 4
    FXMEDITAR15 = 5
    FXMEDITAR22 = 17
    FXMEDITAR30 = 6
    FXMEDITAR34 = 46
    FXMEDITAR38 = 18
    FXMEDITAR42 = 45
    FXMEDITAR46 = 38
    FXMEDITAR50 = 39
    FXMEDITAR54 = 40
    FXMEDITARLIBERADO = 44
End Enum
'CHOTS | Constantes de Meditaciones
'CHOTS | Indican a partir de q lvl se usa tal FX
'31/10/10
'Le agregue la meditacion de los liberados


Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

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
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
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
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
End Enum

Public Const DRAGON As Integer = 6

Public Const MAX_MENSAJES_FORO As Byte = 35

Public Const MAXUSERHECHIZOS As Byte = 25


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const FX_TELEPORT_INDEX As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAXREP As Long = 6000000
Public Const MAXORO As Double = 999999999
Public Const MAXEXP As Double = 1999999999

Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58
Public Const Chala As Integer = 557
Public Const PielLobo As Integer = 414
Public Const PielOsoPardo As Integer = 415
Public Const PielOsoPolar As Integer = 416

Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Integer = 127
Public Const TIJERA_DRUIDA As Integer = 804
Public Const PIQUETE_MINERO As Integer = 187

Public Const DAGA As Integer = 15
Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63
Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192
Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const ObjArboles As Integer = 4
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138
Public Const HILO_SASTRE = 805
Public Const OLLA = 556

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Timbero = 7
    Guardiascaos = 8
    'CHOTS | Nuevos Npcs
    reydelcastillo = 9
    Gobernador = 10
    Quester = 11
    Pasajes = 12
    Cirujano = 13
    Puntos = 14
    Secuas = 15
    Ermitano = 16
    Trader = 17
    Duelero = 18
    OrganizaGuerras = 19
    'CHOTS | Nuevos Npcs
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 24

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 17

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 6


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = 777

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum
''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
    Apuñalar = 7
    Ocultarse = 8
    Supervivencia = 9
    Talar = 10
    Comerciar = 11
    Defensa = 12
    Pesca = 13
    Mineria = 14
    Carpinteria = 15
    Herreria = 16
    Liderazgo = 17
    Domar = 18
    Proyectiles = 19
    Wresterling = 20
    Navegacion = 21
    Alquimia = 22
    Sastreria = 23
    Botanica = 24
End Enum

Public Const FundirMetal = 88
Public Const CapturarNpc = 89 'CHOTS | Monturas

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6
Public Const SND_CAPTURA As Byte = 106 'CHOTS | Monturas
Public Const SND_ROTURA As Byte = 1 'CHOTS | Monturas

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_BEBER As Byte = 46

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 20
Public Const MAX_INVENTORY_SLOTS_NPC As Byte = 50 'CHOTS | Maximos slots

' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otCONTENEDORES = 7
    otCARTELES = 8
    otLlaves = 9
    otFOROS = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otHerramientas = 18
    otTELEPORT = 19
    otMuebles = 20
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otMANCHAS = 35          'No se usa
    otPasajes = 36
    otBandera = 37
    otGemaCaptura = 38 'CHOTS | Monturas
    otMontura = 39 'CHOTS | Monturas
    otCualquiera = 1000
End Enum

'Texto
'CHOTS | Modificado por CHOTS para optimizar paquetes
Public Const FONTTYPE_TALK As String = "~1" ' "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT As String = "~2" ' "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~3" ' "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~4" ' "~65~190~156~0~0"
Public Const FONTTYPE_GEMA As String = "~5" ' "~255~0~255~1~0"
Public Const FONTTYPE_APU As String = "~6" ' "~255~128~0~1~0"
Public Const FONTTYPE_DIOS As String = "~7" ' "~0~240~0~1~0"
Public Const FONTTYPE_SEMI As String = "~8" ' "~255~255~128~1~0"
Public Const FONTTYPE_INFON As String = "~9" ' "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~10" ' "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~11" ' "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~12" ' "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~13" ' "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~14" ' "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~15" ' "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~16" ' "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~17" ' "~255~60~0~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~18" ' "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~19" ' "~255~50~0~1~0"
Public Const FONTTYPE_ORO As String = "~20" ' "~255~255~0~1~0"
Public Const FONTTYPE_CELESTE_NEGRITA As String = "~21" ' "~0~128~255~1~0"
Public Const FONTTYPE_AZUL As String = "~22" ' "~0~0~255~1~0"
Public Const FONTTYPE_GM As String = "~23" ' "~0~0~255~1~0"
Public Const FONTTYPE_TROFORO As String = "~24" ' "~0~0~255~1~0"
Public Const FONTTYPE_TROFPLATA As String = "~25" ' "~0~0~255~1~0"
Public Const FONTTYPE_CELESTE As String = "~26" ' "~255~55~155~1~0"
Public Const FONTTYPE_DUELO As String = "~27" ' "~128~64~64~1~0"
Public Const FONTTYPE_HOGAR As String = "~28" ' "~128~64~64~1~0"
Public Const FONTTYPE_INVOCACION As String = "~29" ' "~128~64~64~1~0"
Public Const FONTTYPE_TORNEOAUTO As String = "~30" ' "~0~74~149~1~0"
Public Const FONTTYPE_MONTURA As String = "~31" ' "~0~2~134~45~0"
Public Const FONTTYPE_GUERRA As String = "~32" ' "~235~235~188~1~0"


'Estadisticas
Public Const STAT_MAXELV As Byte = 54
Public Const STAT_MAXHP As Integer = 999
Public Const STAT_MAXSTA As Integer = 999
Public Const STAT_MAXMAN As Integer = 4000
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99
Public Const STAT_MAXATRIBUTOS As Byte = 36
Public Const STAT_MAXATRIBUTOS_HUMANO As Byte = 35


' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Tipo As TipoHechizo
    
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
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Invoca As Byte
    MinLevel As Byte 'CHOTS | Level para Hechizos
    numNpc As Integer
    Cant As Integer
    
    Materializa As Byte
    itemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean
End Type


Public Type LevelSkill
    LevelValue As Integer
End Type

    Public Type UserOBJ
        ObjIndex As Integer
        Amount As Integer
        Equipped As Byte
        ProbTirar As Byte
    End Type

Public Type Inventario
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
    MonturaObjIndex As Integer 'CHOTS | Monturas
    MonturaSlot As Integer 'CHOTS | Monturas
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
End Type

Public Type InventarioNpc 'CHOTS | Inventario NPC
    Object(1 To MAX_INVENTORY_SLOTS_NPC) As UserOBJ
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
    NroItems As Integer
End Type

Public Type tPartyData
    PIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    TargetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    mapa As Integer
    X As Integer
    Y As Integer
    Name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Apuñala As Byte
    Pegadoble As Byte
    DosManos As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
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
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    Bandera As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    Chala As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    SkSastreria As Integer
    SkAlquimia As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    Jerarquia As Integer 'CHOTS | Jerarquía de los items
    
    NoSeCae As Integer
    
    StaffPower As Integer
    VaraDragon As Byte
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    'CHOTS | Monturas
    MinAta As Integer
    MaxAta As Integer
    MinDefensa As Integer
    MaxDefensa As Integer
    MinMag As Integer
    MaxMag As Integer
    MinAtaMag As Integer
    MaxAtaMag As Integer
    BodyMontura As Integer
    'CHOTS | Monturas
    
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
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

Public Type tReputacion 'Fama del usuario
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double
End Type

'Estadisticas de los usuarios
Public Type UserStats
    TrofOro As Byte
    TorneosAuto As Byte 'CHOTS | Trofeos Automaticos
    TrofBronce As Byte
    TrofPlata As Byte
    DuelosConsec As Byte 'CHOTS | Duelos
    GLD As Long 'Dinero
    Banco As Long
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
        
    def As Integer
    Exp As Double
    ELV As Long
    ELU As Double
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
End Type

'CHOTS | Sistema de Quests
Public Type UserQuest
    numNpc As Integer
    cantNpc As Integer
    nroQuest As Byte
    cantMatados As Integer
End Type
'CHOTS | Sistema de Quests

Public Type UserRetos
    TipoReto As Byte
    Pareja As String
    EsperandoReto As Boolean
    Oponente As String
    enReto As Boolean
    TimeReto As Byte
    GanoReto As Boolean
    PerdioReto As Boolean
    AceptoReto As Boolean
    priceGold As Long
    pricePuntos As Long
    priceItems As Byte
    EnvioRequest As Boolean
End Type

'Flags
Public Type UserFlags
    EstaEmpo As Byte    'Empollando (by yb)
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    Descuento As String
    Casado As Byte
    Ofrecio As Byte
    Hambre As Byte
    ParejaDuelo As Integer 'CHOTS | Duelos 2vs2
    Sed As Byte
    YaFundoClan As Byte
    enTorneoAuto As Boolean 'CHOTS | Torneos automáticos
    enDueloTorneoAuto As Boolean 'CHOTS | Torneos automáticos
    enDuelo As Boolean 'CHOTS | Duelos
    montandoMascota As Byte 'CHOTS | Monturas
    Liberado As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Marcado As Byte 'CHOTS | Marcas
    YaDenuncio As Byte
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
    Navegando As Byte
    Seguro As Boolean
    SeguroClan As Boolean
    SeguroResu As Boolean 'CHOTS | Seguro de Resu
    SeguroCaos As Boolean 'CHOTS | Seguro de Caos
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte
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
    Privilegios As PlayerType
    EsRolesMaster As Boolean
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    
    '[el oso]
    MD5Reportado As String
    '[/el oso]
    
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    NoActualizado As Boolean
    PertAlCons As Byte
    PertAlConsCaos As Byte
    
    Silenciado As Byte
    
    Mimetizado As Byte
End Type

Public Type UserCounters
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
    Hogar As Integer 'CHOTS | /HOGAR
    Torneo As Integer 'CHOTS | Torneos Automáticos
    Estupidez As Integer
    Invisibilidad As Integer
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerUsarFlechas As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
End Type


'CHOTS | Reprogramado todo lo de facciones
Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    Jerarquia As Byte
    RecibioExpInicial As Byte
    RecibioArmadura As Byte
    Reenlistadas As Byte
    FueCaos As Byte
    FueReal As Byte
    Amatar As Integer
End Type

'CHOTS | Guerras
Public Type NPCGuerra
    enGuerra As Boolean
    team As Byte
End Type

Public Type UserGuerra
    enGuerra As Boolean
    status As Byte
    team As Byte
    Sala As Byte
    OldInvent As Inventario
End Type

'Tipo de los Usuarios
Public Type User
    Name As String
    Pareja As String
    Puntos As Integer 'CHOTS | Puntos de usuario
    ID As Long
    
    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    modName As String
    Password As String
    
    char As char 'Define la apariencia
    CharMimetizado As char
    OrigChar As char
    
    Desc As String ' Descripcion
    DescRM As String
    Clase As String
    Raza As String
    Genero As String
    email As String
    Hogar As String
    Preg As String
    Resp As String
        
    Invent As Inventario
    
    Pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
    ColaSalida As New Collection
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    Quest As UserQuest 'CHOTS | Sistema de Quests
    Reto As UserRetos
    
    Reputacion As tReputacion
    
    Faccion As tFacciones

    RandomCode As String
    'CHOTS | Seguridad by DyE
    ClavePublica As Integer
    ClavePrivada As Integer
    'CHOTS | Seguridad by DyE
    
    UseNum As Byte
    UseAcum As Integer
    
    ip As String
    
     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
    
    EmpoCont As Byte
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito

    AreasInfo As AreaInfo

    'CHOTS | Guerras
    guerra As UserGuerra
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

Public Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    UsuariosMatados As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    PuedeMoverse As Boolean
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte

    ExpCount As Long '[ALEJO]
    '[/KEVIN]
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    AtacaAPJ As Integer
    AtacaANPC As Integer
    AIAlineacion As e_Alineacion
    AIPersonalidad As e_Personalidad
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
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
' New type for holding the pathfinding info



Public Type Npc
    Name As String
    char As char 'Define como se vera
    Desc As String
    DescExtra As String

    NPCtype As eNPCType
    Numero As Integer
    
    Level As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte

    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    GiveExp As Double
    GiveGLD As Long

    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As InventarioNpc 'CHOTS | Inventario NPC
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
    
    'CHOTS | Monturas
    probCaptura As Byte

    'CHOTS | Guerras
    guerra As NPCGuerra
    salaGuerra As Byte
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    
    Terreno As String
    zona As String
    Restringir As String
    BackUp As Byte

    MinLevel As Byte
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public ULTIMAVERSION As String
Public BackUp As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public Torneo_Clases_Validas(1 To 8) As String
Public Torneo_Alineacion_Validas(1 To 8) As String
Public Torneo_Clases_Validas2(1 To 8) As Integer
Public Torneo_Alineacion_Validas2(1 To 4) As Integer
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Const ENDL As String * 2 = vbCrLf
Public Const ENDC As String * 1 = vbNullChar

Public MultExp As Byte
Public MultOro As Byte
Public MultHogar As Byte
Public MinTrabajo As Byte
Public MaxTrabajo As Byte

Public recordusuarios As Long

'CHOTS | EscucharClan
Public Clan_EscuchadorIndex As Integer
Public Clan_ClanIndex As Integer
'CHOTS | EscucharClan

'CHOTS | Espia de Users
Public Espia_Espiador As Integer
Public Espia_Espiado As Integer
'CHOTS | Espia de Users


'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos ' TODO: Se usa esta variable ?

''
'Numero de usuarios actual
Public NumUsers As Integer
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
Public Torneo_SumAuto As Integer
Public Torneo_Map As Integer
Public Torneo_X As Integer
Public Torneo_Y As Integer
Public Hay_Torneo As Boolean
Public Torneo_Nivel_Minimo As Long
Public Torneo_Nivel_Maximo As Long
Public Torneo_Cantidad As Long
Public Torneo_Inscriptos As Long
Public Oscuridad As Integer
Public NocheDia As Integer
Public PuedeCrearPersonajes As Integer
Public CamaraLenta As Integer
Public ServerSoloGMs As Integer

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte


Public EnPausa As Boolean
Public EncriptarProtocolosCriticos As Boolean


'*****************ARRAYS PUBLICOS*************************
Public ArrayIps() As String
Public UserList() As User 'USUARIOS
Public Npclist() As Npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To STAT_MAXELV) As LevelSkill
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public ObjDruida() As Integer
Public ObjSastre() As Integer
Public MD5s() As String
Public BanIps As New Collection
Public Parties() As clsParty
'*********************************************************

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As New cCola
Public Torneo As New cCola
Public ConsultaPopular As New ConsultasPopulares
Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

'Bysnack | Mensajes genericos de Gema
Public Const HP As String = "20 de vida extra."
Public Const MP As String = "100 de mana extra."
Public Const EV As String = "Aumento de la evasión."
Public Const PA As String = "Aumento de la probabilidad de apuñalar."
Public Const PG As String = "Aumento de la probabilidad de golpe."
Public Const DE As String = "Mayor defensa con escudos."

'CHOTS | Cada cuanto va el centinela
Public Const MAX_TRABAJO_CENTINELA = 800
