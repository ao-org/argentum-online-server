Attribute VB_Name = "Protocol"

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR             As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 255

Public Enum ServerPacketID

    logged                  ' LOGGED  0
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU   10
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    CharSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF 20
    PartySafeOn
    PartySafeOff
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE 30
    changeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+   40
    ShowMessageBox          ' !!
    MostrarCuenta
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    fxpiso
    ObjectDelete            ' BO  50
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01 60
    ChangeInventorySlot     ' CSI
    InventoryUnlockSlots
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU 70
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER 80
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR 90
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    FYA
    CerrarleCliente
    Contadores
    
    'GM messages
    SpawnListt               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    PersonajesDeCuenta
    UserOnline '110
    ParticleFX
    ParticleFXToFloor
    ParticleFXWithDestino
    ParticleFXWithDestinoXY
    Hora
    light
    AuraToChar
    SpeedToChar
    LightToFloor
    NieveToggle
    NieblaToggle
    Goliath
    TextOverChar
    TextOverTile
    TextCharDrop
    FlashScreen
    AlquimistaObj
    ShowAlquimiaForm
    Familiar
    SastreObj
    ShowSastreForm ' 126
    VelocidadToggle
    MacroTrabajoToggle
    RefreshAllInventorySlot
    BindKeys
    ShowFrmLogear
    ShowFrmMapa
    InmovilizadoOK
    BarFx
    SetEscribiendo
    Logros
    TrofeoToggleOn
    TrofeoToggleOff
    LocaleMsg
    ListaCorreo
    ShowPregunta
    DatosGrupo
    ubicacion
    CorreoPicOn
    DonadorObj
    ArmaMov
    EscudoMov
    ActShop
    ViajarForm
    Oxigeno
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    Ranking
    PosLLamadaDeClan
    QuestDetails
    QuestListSend
    NpcQuestListSend
    UpdateNPCSimbolo
    ClanSeguro
    Intervals
    UpdateUserKey
    UpdateRM
    UpdateDM
    RequestProcesses
    RequestScreenShot
    ShowProcesses
    ShowScreenShot
    ScreenShotData
    Tolerancia0
    Redundancia
    SeguroResu
    Stopped
    InvasionInfo
    CommerceRecieveChatMessage

End Enum

Private Enum ClientPacketID

    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    ThrowDice
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    PartySafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    PetLeave                '/LIBERAR
    GrupoMsg                '/GrupoMsg
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    punishments             '/PENAS
    ChangePassword          '/Contraseña
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    Ping                    '/PING
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    banip                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    Participar           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    RequestTCPStats         '/TCPESSTATS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    'Nuevas Ladder
    GlobalMessage           '/CONSOLA
    GlobalOnOff
    IngresarConCuenta
    BorrarPJ
    NewPacketID
    Desbuggear
    DarLlaveAUsuario
    SacarLlave
    VerLlaves
    UseKey
    Day
    SetTime
    DonateGold              '/DONAR
    Promedio                '/PROMEDIO
    GiveItem                '/DAR

End Enum

Private Enum NewPacksID

    OfertaInicial
    OfertaDeSubasta
    QuestionGM
    CuentaRegresiva
    PossUser
    Duel
    AcceptDuel
    CancelDuel
    QuitDuel
    NieveToggle
    NieblaToggle
    TransFerGold
    Moveitem
    Genio
    Casarse
    CraftAlquimista
    RequestFamiliar
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    Escribiendo
    TraerRecompensas
    ReclamarRecompensa
    Correo
    SendCorreo
    RetirarItemCorreo
    BorrarCorreo
    InvitarGrupo
    ResponderPregunta
    RequestGrupo
    AbandonarGrupo
    HecharDeGrupo
    MacroPossent
    SubastaInfo
    BanCuenta
    unBanCuenta
    BanSerial
    unBanSerial
    CerrarCliente
    EventoInfo
    CrearEvento
    BanTemporal
    Traershop
    ComprarItem
    ScrollInfo
    CancelarExit
    EnviarCodigo
    CrearTorneo
    ComenzarTorneo
    CancelarTorneo
    BusquedaTesoro
    CompletarViaje
    BovedaMoveItem
    QuieroFundarClan
    LlamadadeClan
    MarcaDeClanPack
    MarcaDeGMPack
    TraerRanking
    Pareja
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    SeguroClan
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    Home                    '/HOGAR
    Consulta                '/CONSULTA
    RequestScreenShot       '/SS
    RequestProcesses        '/VERPROCESOS
    SendScreenShot
    SendProcesses
    Tolerancia0             '/T0
    GetMapInfo              '/MAPINFO
    FinEvento
    SeguroResu
    CuentaExtractItem
    CuentaDeposit
    CreateEvent
    CommerceSendChatMessage
    LogMacroClickHechizo

End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza
    eo_Arma
    eo_Escudo
    eo_Casco
    eo_Particula
    eo_Vida
    eo_Mana
    eo_Energia
    eo_MinHP
    eo_MinMP
    eo_Hit
    eo_MinHit
    eo_MaxHit
    eo_Desc
    eo_Intervalo
    eo_Hogar

End Enum

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_DIOS
    FONTTYPE_CITIZEN
    FONTTYPE_CRIMINAL
    FONTTYPE_EXP
    FONTTYPE_SUBASTA
    FONTTYPE_GLOBAL
    FONTTYPE_MP
    FONTTYPE_ROSA
    FONTTYPE_VIOLETA
    FONTTYPE_INFOIAO
    
    FONTTYPE_New_Amarillo_Oscuro
    FONTTYPE_New_Verde_Oscuro
    FONTTYPE_New_Naranja
    FONTTYPE_New_Celeste
    FONTTYPE_New_Amarillo_Verdoso
    FONTTYPE_New_Gris
    FONTTYPE_New_Blanco
    FONTTYPE_New_Rojo_Salmon
    FONTTYPE_New_DONADOR
    FONTTYPE_New_GRUPO
    FONTTYPE_New_Eventos
    
    FONTTYPE_PROMEDIO_IGUAL
    FONTTYPE_PROMEDIO_MENOR
    FONTTYPE_PROMEDIO_MAYOR
    
End Enum

Public Type PersonajeCuenta

    nombre As String
    nivel As Byte
    Mapa As Integer
    posX As Integer
    posY As Integer
    cuerpo As Integer
    Cabeza As Integer
    Status As Byte
    clase As Byte
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    ClanIndex As Integer

End Type

Public Type t_DataBuffer
    data() As Byte
    Length As Integer
End Type

''
' Handles incoming data.
'
' @param    UserIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)
        
        ' [2020-5-23 Mateo] Esto es normal que suceda, puede existir un paquete INCOMPLETO y esto hace que no lo procese y deje acumulado el buffer para el proximo dato
        If Not .incomingData.CheckLength Then
            Debug.Print "Not .IncomingData.CheckLength! Último paquete: " & .LastPacketID & IIf(.LastPacketID = ClientPacketID.NewPacketID, " (New: " & .LastNewPacketID & ") -", " - ") & Date$ & " - " & Time$; ""
            HandleIncomingData = False
            Exit Function
        End If
    
        If Not .incomingData.ValidCRC Then
            Debug.Print "UserIndex: " & UserIndex & " El paquete es invalido, posible hack, echarlo!"
            HandleIncomingData = False
            Call CloseSocket(UserIndex)
            'Stop
            Exit Function
        End If
    
        Dim PacketID As Long
            PacketID = CLng(.incomingData.ReadID())
    
        'Does the packet requires a logged user??
        If Not (PacketID = ClientPacketID.LoginExistingChar Or _
                PacketID = ClientPacketID.LoginNewChar Or _
                PacketID = ClientPacketID.IngresarConCuenta Or _
                PacketID = ClientPacketID.BorrarPJ Or _
                PacketID = ClientPacketID.ThrowDice) Then
            
            'Is the user actually logged?
            If Not .flags.UserLogged Then
                Call CloseSocket(UserIndex)
                Exit Function
            
                'He is logged. Reset idle counter if id is valid.
            ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
                .Counters.IdleCount = 0
    
            End If
    
        Else
        
            .Counters.IdleCount = 0
            
            ' Envió el primer paquete
            .flags.FirstPacket = True
                
            #If AntiExternos = 1 Then
                .Redundance = RandomNumber(2, 255)
                Call WriteRedundancia(UserIndex)
            #End If
    
        End If

    End With
    
    Select Case PacketID
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
    
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
    
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)

        Case ClientPacketID.IngresarConCuenta
            Call HandleIngresarConCuenta(UserIndex)

        Case ClientPacketID.BorrarPJ
            Call HandleBorrarPJ(UserIndex)

        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)
            
        Case ClientPacketID.ThrowDice
            Call HandleThrowDice(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.PartySafeToggle
            Call HandlePartyToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
           
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
            
        Case ClientPacketID.PetLeave                '/LIBERAR
            Call HandlePetLeave(UserIndex)
        
        Case ClientPacketID.GrupoMsg
            Call HandleGrupoMsg(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
                
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)

        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.ChangePassword          '/Contraseña
            Call HandleChangePassword(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
        
            'GM messages
        Case ClientPacketID.GMMessage               '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case ClientPacketID.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case ClientPacketID.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case ClientPacketID.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case ClientPacketID.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case ClientPacketID.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case ClientPacketID.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case ClientPacketID.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case ClientPacketID.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case ClientPacketID.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case ClientPacketID.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case ClientPacketID.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
        
        Case ClientPacketID.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case ClientPacketID.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
            
        Case ClientPacketID.Desbuggear              '/DESBUGGEAR
            Call HandleDesbuggear(UserIndex)
            
        Case ClientPacketID.DarLlaveAUsuario        '/DARLLAVE
            Call HandleDarLlaveAUsuario(UserIndex)
            
        Case ClientPacketID.SacarLlave              '/SACARLLAVE
            Call HandleSacarLlave(UserIndex)
            
        Case ClientPacketID.VerLlaves               '/VERLLAVES
            Call HandleVerLlaves(UserIndex)
            
        Case ClientPacketID.UseKey
            Call HandleUseKey(UserIndex)
        
        Case ClientPacketID.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case ClientPacketID.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)
        
        Case ClientPacketID.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
            
        Case ClientPacketID.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case ClientPacketID.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case ClientPacketID.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
        
        Case ClientPacketID.Forgive                 '/PERDON
            Call HandleForgive(UserIndex)
            
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
            
        Case ClientPacketID.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
            
        Case ClientPacketID.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(UserIndex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
            
        Case ClientPacketID.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)
            
        Case ClientPacketID.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
            
        Case ClientPacketID.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
            
        Case ClientPacketID.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
            
        Case ClientPacketID.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(UserIndex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(UserIndex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(UserIndex)
        
        Case ClientPacketID.banip                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
        
        Case ClientPacketID.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)
        
        Case ClientPacketID.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case ClientPacketID.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(UserIndex)
        
        Case ClientPacketID.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(UserIndex)
        
        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case ClientPacketID.Participar           '/APAGAR
            Call HandleParticipar(UserIndex)
        
        Case ClientPacketID.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(UserIndex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(UserIndex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case ClientPacketID.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)
        
        Case ClientPacketID.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(UserIndex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)
            
        Case ClientPacketID.night                   '/NOCHE
            Call HandleNight(UserIndex)

        Case ClientPacketID.Day                     '/DIA
            Call HandleDay(UserIndex)

        Case ClientPacketID.SetTime                 '/HORA X
            Call HandleSetTime(UserIndex)

        Case ClientPacketID.DonateGold              '/DONAR
            Call HandleDonateGold(UserIndex)
                
        Case ClientPacketID.Promedio                '/PROMEDIO
            Call HandlePromedio(UserIndex)
                
        Case ClientPacketID.GiveItem                '/DAR
            Call HandleGiveItem(UserIndex)

        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case ClientPacketID.RequestTCPStats         '/TCPESSTATS
            Call HandleRequestTCPStats(UserIndex)
        
        Case ClientPacketID.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case ClientPacketID.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case ClientPacketID.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case ClientPacketID.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case ClientPacketID.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)
        
        Case ClientPacketID.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(UserIndex)
        
        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
            
            'Nuevo Ladder
            
        Case ClientPacketID.GlobalMessage           '/CONSOLA
            Call HandleGlobalMessage(UserIndex)
        
        Case ClientPacketID.GlobalOnOff             '/GLOBAL
            Call HandleGlobalOnOff(UserIndex)
        
        Case ClientPacketID.NewPacketID    'Los Nuevos Packs ID
            Call HandleIncomingDataNewPacks(UserIndex)

        Case Else
            Call RegistrarError(-1, "Paquete inválido: " & PacketID & " UserIndex: " & UserIndex & " (IP: " & UserList(UserIndex).ip & ") Último paquete: " & UserList(UserIndex).LastPacketID & IIf(UserList(UserIndex).LastPacketID = ClientPacketID.NewPacketID, " (New: " & UserList(UserIndex).LastNewPacketID & ")", ""), "Protocol.HandleIncomingData", Erl)
            Call CloseSocket(UserIndex)

    End Select
    
    With UserList(UserIndex).incomingData
    
        Call .ReadNewPacket

        'Done with this packet, move on to next one or send everything if no more packets found
        If (Not .BufferOver Or .Length > 0) And .errNumber = 0 Then
            Call Err.Clear
            HandleIncomingData = True
      
        ElseIf .errNumber <> 0 And .errNumber <> .NotEnoughDataErrCode Then
            'An error ocurred, log it and kick player.
            Call RegistrarError(Err.Number, Err.Description & vbNewLine & "PackedId: " & PacketID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "UserName: " & UserList(UserIndex).name, "UserIndex: " & UserIndex), "Protocol.HandleIncomingData", Erl)
            
            Call CloseSocket(UserIndex)
      
            HandleIncomingData = False
            
        Else
        
            HandleIncomingData = False
    
        End If
        
        .errNumber = 0
        UserList(UserIndex).LastPacketID = PacketID
    
    End With
    

End Function

Public Sub HandleIncomingDataNewPacks(ByVal UserIndex As Integer)

    '***************************************************
    'Los nuevos Pack ID
    'Creado por Ladder con gran ayuda de Maraxus
    '04.12.08
    '***************************************************
    Dim PacketID As Integer
        PacketID = UserList(UserIndex).incomingData.ReadByte

    Select Case PacketID

        Case NewPacksID.OfertaInicial
            Call HandleOfertaInicial(UserIndex)
    
        Case NewPacksID.OfertaDeSubasta
            Call HandleOfertaDeSubasta(UserIndex)
        
        Case NewPacksID.CuentaRegresiva
            Call HandleCuentaRegresiva(UserIndex)

        Case NewPacksID.QuestionGM
            Call HandleQuestionGM(UserIndex)

        Case NewPacksID.PossUser
            Call HandlePossUser(UserIndex)

        Case NewPacksID.Duel
            Call HandleDuel(UserIndex)
                
        Case NewPacksID.AcceptDuel
            Call HandleAcceptDuel(UserIndex)
                
        Case NewPacksID.CancelDuel
            Call HandleCancelDuel(UserIndex)
                
        Case NewPacksID.QuitDuel
            Call HandleQuitDuel(UserIndex)

        Case NewPacksID.NieveToggle
            Call HandleNieveToggle(UserIndex)

        Case NewPacksID.NieblaToggle
            Call HandleNieblaToggle(UserIndex)

        Case NewPacksID.TransFerGold
            Call HandleTransFerGold(UserIndex)

        Case NewPacksID.Moveitem
            Call HandleMoveItem(UserIndex)

        Case NewPacksID.LlamadadeClan
            Call HandleLlamadadeClan(UserIndex)

        Case NewPacksID.QuieroFundarClan
            Call HandleQuieroFundarClan(UserIndex)

        Case NewPacksID.BovedaMoveItem
            Call HandleBovedaMoveItem(UserIndex)

        Case NewPacksID.Genio
            Call HandleGenio(UserIndex)

        Case NewPacksID.Casarse
            Call HandleCasamiento(UserIndex)

        Case NewPacksID.EnviarCodigo
            Call HandleEnviarCodigo(UserIndex)

        Case NewPacksID.CrearTorneo
            Call HandleCrearTorneo(UserIndex)
            
        Case NewPacksID.ComenzarTorneo
            Call HandleComenzarTorneo(UserIndex)
            
        Case NewPacksID.CancelarTorneo
            Call HandleCancelarTorneo(UserIndex)

        Case NewPacksID.BusquedaTesoro
            Call HandleBusquedaTesoro(UserIndex)

        Case NewPacksID.CraftAlquimista
            Call HandleCraftAlquimia(UserIndex)

        Case NewPacksID.RequestFamiliar
            Call HandleRequestFamiliar(UserIndex)

        Case NewPacksID.FlagTrabajar
            Call HandleFlagTrabajar(UserIndex)

        Case NewPacksID.CraftSastre
            Call HandleCraftSastre(UserIndex)

        Case NewPacksID.MensajeUser
            Call HandleMensajeUser(UserIndex)

        Case NewPacksID.TraerBoveda
            Call HandleTraerBoveda(UserIndex)

        Case NewPacksID.CompletarAccion
            Call HandleCompletarAccion(UserIndex)

        Case NewPacksID.Escribiendo
            Call HandleEscribiendo(UserIndex)

        Case NewPacksID.TraerRecompensas
            Call HandleTraerRecompensas(UserIndex)

        Case NewPacksID.ReclamarRecompensa
            Call HandleReclamarRecompensa(UserIndex)

        Case NewPacksID.Correo
            Call HandleCorreo(UserIndex)

        Case NewPacksID.SendCorreo ' ok
            Call HandleSendCorreo(UserIndex)

        Case NewPacksID.RetirarItemCorreo ' ok
            Call HandleRetirarItemCorreo(UserIndex)

        Case NewPacksID.BorrarCorreo
            Call HandleBorrarCorreo(UserIndex) 'ok

        Case NewPacksID.InvitarGrupo
            Call HandleInvitarGrupo(UserIndex) 'ok

        Case NewPacksID.MarcaDeClanPack
            Call HandleMarcaDeClan(UserIndex)

        Case NewPacksID.MarcaDeGMPack
            Call HandleMarcaDeGM(UserIndex)

        Case NewPacksID.ResponderPregunta 'ok
            Call HandleResponderPregunta(UserIndex)

        Case NewPacksID.RequestGrupo
            Call HandleRequestGrupo(UserIndex) 'ok

        Case NewPacksID.AbandonarGrupo
            Call HandleAbandonarGrupo(UserIndex) ' ok

        Case NewPacksID.HecharDeGrupo
            Call HandleHecharDeGrupo(UserIndex) 'ok

        Case NewPacksID.MacroPossent
            Call HandleMacroPos(UserIndex)

        Case NewPacksID.SubastaInfo
            Call HandleSubastaInfo(UserIndex)

        Case NewPacksID.EventoInfo
            Call HandleEventoInfo(UserIndex)

        Case NewPacksID.CrearEvento
            Call HandleCrearEvento(UserIndex)

        Case NewPacksID.BanCuenta
            Call HandleBanCuenta(UserIndex)
            
        Case NewPacksID.unBanCuenta
            Call HandleUnBanCuenta(UserIndex)
            
        Case NewPacksID.BanSerial
            Call HandleBanSerial(UserIndex)
        
        Case NewPacksID.unBanSerial
            Call HandleUnBanSerial(UserIndex)
            
        Case NewPacksID.CerrarCliente
            Call HandleCerrarCliente(UserIndex)
            
        Case NewPacksID.BanTemporal
            Call HandleBanTemporal(UserIndex)

        Case NewPacksID.Traershop
            Call HandleTraerShop(UserIndex)

        Case NewPacksID.TraerRanking
            Call HandleTraerRanking(UserIndex)

        Case NewPacksID.Pareja
            Call UserList(UserIndex).incomingData.ReadInteger ' Desactivado. Nada para hacer
            
        Case NewPacksID.ComprarItem
            Call HandleComprarItem(UserIndex)
            
        Case NewPacksID.CompletarViaje
            Call HandleCompletarViaje(UserIndex)
            
        Case NewPacksID.ScrollInfo
            Call HandleScrollInfo(UserIndex)

        Case NewPacksID.CancelarExit
            Call HandleCancelarExit(UserIndex)
            
        Case NewPacksID.Quest
            Call HandleQuest(UserIndex)
            
        Case NewPacksID.QuestAccept
            Call HandleQuestAccept(UserIndex)
        
        Case NewPacksID.QuestListRequest
            Call HandleQuestListRequest(UserIndex)
        
        Case NewPacksID.QuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)
        
        Case NewPacksID.QuestAbandon
            Call HandleQuestAbandon(UserIndex)
            
        Case NewPacksID.SeguroClan
            Call HandleSeguroClan(UserIndex)
            
        Case NewPacksID.CreatePretorianClan     '/CREARPRETORIANOS
            Call HandleCreatePretorianClan(UserIndex)
         
        Case NewPacksID.RemovePretorianClan     '/ELIMINARPRETORIANOS
            Call HandleDeletePretorianClan(UserIndex)

        Case NewPacksID.Home
            Call HandleHome(UserIndex)
            
        Case NewPacksID.Consulta
            Call HandleConsulta(UserIndex)
                
        Case NewPacksID.RequestScreenShot       '/SS
            Call HandleRequestScreenShot(UserIndex)
                
        Case NewPacksID.RequestProcesses        '/VERPROCESOS
            Call HandleRequestProcesses(UserIndex)
                
        Case NewPacksID.Tolerancia0             '/T0
            Call HandleTolerancia0(UserIndex)

        Case NewPacksID.GetMapInfo
            Call HandleGetMapInfo(UserIndex)
                
        Case NewPacksID.FinEvento
            Call HandleFinEvento(UserIndex)
                
        Case NewPacksID.SendScreenShot
            Call HandleScreenShot(UserIndex)
                
        Case NewPacksID.SendProcesses
            Call HandleProcesses(UserIndex)

        Case NewPacksID.SeguroResu
            Call HandleSeguroResu(UserIndex)

        Case NewPacksID.CuentaExtractItem
            Call HandleCuentaExtractItem(UserIndex)
                
        Case NewPacksID.CuentaDeposit
            Call HandleCuentaDeposit(UserIndex)
                
        Case NewPacksID.CreateEvent
            Call HandleCreateEvent(UserIndex)
                
        Case NewPacksID.CommerceSendChatMessage
            Call HandleCommerceSendChatMessage(UserIndex)
                
        Case NewPacksID.LogMacroClickHechizo
            Call HandleLogMacroClickHechizo(UserIndex)
            
        Case Else
            Call RegistrarError(-1, "New paquete inválido: " & PacketID & " UserIndex: " & UserIndex & " (IP: " & UserList(UserIndex).ip & ")", "Protocol.HandleIncomingDataNewPacks", Erl)
            Call CloseSocket(UserIndex)
            
    End Select
    
    UserList(UserIndex).LastNewPacketID = PacketID
        
End Sub

Public Function ConvertDataBuffer(ByVal Length As Integer, _
                                  ByRef data() As Byte) As t_DataBuffer
    
    ConvertDataBuffer.data = data
    ConvertDataBuffer.Length = Length
    
End Function

''
' Handles the "LoginExistingChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    ''Last Modification: 01/12/08 Ladder
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserName    As String
    Dim CuentaEmail As String
    Dim Password    As String
    Dim Version     As String
    Dim MacAddress  As String
    Dim HDserial    As Long
    Dim MD5         As String
        
    With UserList(UserIndex).incomingData

        CuentaEmail = .ReadASCIIString()
        Password = .ReadASCIIString()
        Version = CStr(.ReadByte()) & "." & CStr(.ReadByte()) & "." & CStr(.ReadByte())
        UserName = .ReadASCIIString()
        MacAddress = .ReadASCIIString()
        HDserial = .ReadLong()
        MD5 = .ReadASCIIString()
        
    End With

    #If DEBUGGING = False Then

        If Not VersionOK(Version) Then
            Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

    #End If
        
    If EsGmChar(UserName) Then
            
        If AdministratorAccounts(UCase$(UserName)) <> UCase$(CuentaEmail) Then
            Call WriteShowMessageBox(UserIndex, "¡ESTE PERSONAJE NO TE PERTENECE!")
            Call SaveBanCuentaDatabase(UserList(UserIndex).AccountId, "Intento de hackeo de personajes ajenos", "El Servidor")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
            
    End If
  
    If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MacAddress, HDserial, MD5) Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If Not AsciiValidos(UserName) Then
        Call WriteShowMessageBox(UserIndex, "Nombre invalido.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteShowMessageBox(UserIndex, "El personaje no existe.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If Not PersonajePerteneceID(UserName, UserList(UserIndex).AccountId) Then
        'Call WriteShowMessageBox(UserIndex, "¡ESTE PERSONAJE NO TE PERTENECE!")
        Call LogHackAttemp("Alguien ha tratado de ingresar con el PJ '" & UserName & "' desde una cuenta ajena ID: " & UserList(UserIndex).AccountId & " desde la IP: " & UserList(UserIndex).ip)
        Call SaveBanCuentaDatabase(UserList(UserIndex).AccountId, "Intento de hackeo de personajes ajenos", "El Servidor")
        Call CloseSocket(UserIndex)
        Exit Sub
            
    End If
    
    If BANCheck(UserName) Then

        Dim LoopC As Integer
        
        For LoopC = 1 To Baneos.Count

            If Baneos(LoopC).name = UCase$(UserName) Then
                Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada a Argentum20 hasta el día " & Format(Baneos(LoopC).FechaLiberacion, "dddddd") & " a las " & Format(Baneos(LoopC).FechaLiberacion, "hh:mm am/pm") & " debido a " & Baneos(LoopC).Causa & " Esta decisión fue tomada por " & Baneos(LoopC).Baneador & ".")
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

        Next LoopC
        
        Dim BanNick     As String
        Dim BaneoMotivo As String

        BanNick = GetVar(CharPath & UserName & ".chr", "BAN", "BannedBy")
        BaneoMotivo = GetVar(CharPath & UserName & ".chr", "BAN", "BanMotivo")

        If LenB(BanNick) = 0 Then BanNick = "*Error en la base de datos*"
        If LenB(BaneoMotivo) = 0 Then BaneoMotivo = "*No se registra el motivo del baneo.*"
        
        Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada al juego debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & ".")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
        
    Call ConnectUser(UserIndex, UserName, CuentaEmail)

    Exit Sub
    
ErrHandler:
        
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserName As String
    Dim race     As eRaza
    Dim gender   As eGenero
    Dim Hogar    As eCiudad
    Dim Class As eClass
    Dim Head        As Integer
    Dim CuentaEmail As String
    Dim Password    As String
    Dim MacAddress  As String
    Dim HDserial    As Long
    Dim MD5         As String
    Dim Version     As String
      
    With UserList(UserIndex).incomingData

        CuentaEmail = .ReadASCIIString()
        Password = .ReadASCIIString()
        Version = CStr(.ReadByte()) & "." & CStr(.ReadByte()) & "." & CStr(.ReadByte())
        UserName = .ReadASCIIString()
        race = .ReadByte()
        gender = .ReadByte()
        Class = .ReadByte()
        Head = .ReadInteger()
        Hogar = .ReadByte()
        MacAddress = .ReadASCIIString()
        HDserial = .ReadLong()
        MD5 = .ReadASCIIString()

    End With
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteShowMessageBox(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteShowMessageBox(UserIndex, "Has creado demasiados personajes.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If ObtenerCantidadDePersonajesByUserIndex(UserIndex) >= MAX_PERSONAJES Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    #If DEBUGGING = False Then

        If Not VersionOK(Version) Then
            Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

    #End If
        
    If EsGmChar(UserName) Then
            
        If AdministratorAccounts(UCase$(UserName)) <> UCase$(CuentaEmail) Then
            Call WriteShowMessageBox(UserIndex, "El nombre de usuario ingresado está siendo ocupado por un miembro del Staff.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
            
    End If
        
    If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MacAddress, HDserial, MD5) Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
            
    If GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountId) >= MAX_PERSONAJES Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If Not ConnectNewUser(UserIndex, UserName, race, gender, Class, Head, CuentaEmail, Hogar) Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
        
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleThrowDice(ByVal UserIndex As Integer)
    
    On Error GoTo HandleThrowDice_Err

    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)

    End With
    
    Call WriteDiceRoll(UserIndex)

    Exit Sub

HandleThrowDice_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleThrowDice", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String
            chat = .incomingData.ReadASCIIString()

        '[Consejeros & GMs]
        If (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "Dijo: " & chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
        
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then

                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call EquiparBarco(UserIndex)

                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
    
                End If

            End If

        End If
       
        If .flags.Silenciado = 1 Then
        
            'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
            
        Else

            If LenB(chat) <> 0 Then
            
                'Analize chat...
                Call Statistics.ParseChat(chat)

                ' WyroX: Foto-denuncias - Push message
                Dim i As Long
                For i = 1 To UBound(.flags.ChatHistory) - 1
                    .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                    
                .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR, UserList(UserIndex).name))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor, UserList(UserIndex).name))

                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTalk", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Yell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String
            chat = .incomingData.ReadASCIIString()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
        
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
            
        Else

            '[Consejeros & GMs]
            If (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios)) Then
                Call LogGM(.name, "Grito: " & chat)
            End If
            
            'I see you....
            If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando = 1 Then
                    
                    'TODO: Revisar con WyroX
                    If .clase = eClass.Pirat Then
                    
                        ' Pierde la apariencia de fragata fantasmal
                        .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
    
                        .Char.ShieldAnim = NingunEscudo
                        .Char.WeaponAnim = NingunArma
                        .Char.CascoAnim = NingunCasco
    
                        Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    
                    End If
    
                Else
    
                    If .flags.invisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                End If

            End If
            
            If .flags.Silenciado = 1 Then
                Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        
                'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Else

                If LenB(chat) <> 0 Then
                    'Analize chat...
                    Call Statistics.ParseChat(chat)

                    ' WyroX: Foto-denuncias - Push message
                    Dim i As Long
                    For i = 1 To UBound(.flags.ChatHistory) - 1
                        .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    
                    .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat

                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed, UserList(UserIndex).name))
               
                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleYell", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Whisper" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat            As String
        Dim targetCharIndex As String
        Dim targetUserIndex As Integer

        targetCharIndex = .incomingData.ReadASCIIString()
        chat = .incomingData.ReadASCIIString()
 
        targetUserIndex = NameIndex(targetCharIndex)

        If targetUserIndex <= 0 Then 'existe el usuario destino?
            Call WriteConsoleMsg(UserIndex, "Usuario offline o inexistente.", FontTypeNames.FONTTYPE_INFO)

        Else
        
            If Not EsGM(UserIndex) And EsGM(targetUserIndex) Then
            
                Call WriteConsoleMsg(UserIndex, "No podes hablar por privado con Game Masters.", FontTypeNames.FONTTYPE_WARNING)

            Else

                If EstaPCarea(UserIndex, targetUserIndex) Then

                    If LenB(chat) <> 0 Then
                    
                        'Analize chat...
                        Call Statistics.ParseChat(chat)

                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Long
                        For i = 1 To UBound(.flags.ChatHistory) - 1
                            .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
                        
                        .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
                        Call SendData(SendTarget.ToSuperioresArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, RGB(157, 226, 20)))
                        
                        Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, RGB(157, 226, 20))
                        Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, RGB(157, 226, 20))
                        'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                        'Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                        Call WritePlayWave(targetUserIndex, FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "[" & .name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                    Call WriteConsoleMsg(targetUserIndex, "[" & .name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                    Call WritePlayWave(targetUserIndex, FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)
                    
                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWhisper", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Walk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo HandleWalk_Err

    Dim Heading As eHeading
    
    With UserList(UserIndex)

        Heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then

            If .flags.Meditando Then
            
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))

            End If
                
            Dim CurrentTick As Long
                CurrentTick = GetTickCount
            
            'Prevent SpeedHack (refactored by WyroX)
            If Not EsGM(UserIndex) Then
                Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single
                ElapsedTimeStep = CurrentTick - .Counters.LastStep
                MinTimeStep = .Intervals.Caminar / .Char.speeding
                DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep
                    
                .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep

                If DeltaStep > 0 Then
                
                    If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                        'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración » Posible uso de SpeedHack del usuario " & .name & ".", FontTypeNames.FONTTYPE_SERVER))
                        Call WritePosUpdate(UserIndex)
                        Exit Sub

                    End If

                Else

                    If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                End If

            End If
            
            'Move user
            If MoveUserChar(UserIndex, Heading) Then
            
                ' Save current step for anti-sh
                .Counters.LastStep = CurrentTick
                
                If UserList(UserIndex).Grupo.EnGrupo = True Then
                    Call CompartirUbicacion(UserIndex)

                End If
    
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                        
                    Call WriteRestOK(UserIndex)
                    'Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "178", FontTypeNames.FONTTYPE_INFO)
    
                End If
                        
                Call CancelExit(UserIndex)
                        
                'Esta usando el /HOGAR, no se puede mover
                If .flags.Traveling = 1 Then
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                    Call WriteConsoleMsg(UserIndex, "Has cancelado el viaje a casa.", FontTypeNames.FONTTYPE_INFO)

                End If

                ' Si no pudo moverse
            Else
                .Counters.LastStep = 0
                Call WritePosUpdate(UserIndex)

            End If

        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                'Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "54", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
            
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                
            If .clase <> eClass.Thief And .clase <> eClass.Bandit Then
            
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando = 1 Then
                        
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
        
                        .Char.ShieldAnim = NingunEscudo
                        .Char.WeaponAnim = NingunArma
                        .Char.CascoAnim = NingunCasco
        
                        Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    
                    End If
    
                Else
    
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                    End If
    
                End If
    
            End If
                
        End If

    End With

    Exit Sub

HandleWalk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWalk", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleRequestPositionUpdate_Err

    Call WritePosUpdate(UserIndex)
  
    Exit Sub

HandleRequestPositionUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlRequestPositionUpdate", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Attack" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
        
    On Error GoTo HandleAttack_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/01/08
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '***************************************************
        
    With UserList(UserIndex)
    
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then

            If ObjData(.Invent.WeaponEqpObjIndex).Proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

        End If
        
        If .Invent.HerramientaEqpObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Para atacar debes desequipar la herramienta.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If
        
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
            
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
        
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
                
            If .flags.Navegando = 1 Then

                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
    
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco
    
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

                End If
    
            Else
    
                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFOIAO)
    
                End If
    
            End If
    
        End If

    End With

    Exit Sub

HandleAttack_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAttack", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "PickUp" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
        
    On Error GoTo HandlePickUp_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Lower rank administrators can't pick up items
        If (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Call WriteConsoleMsg(UserIndex, "No podés tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call GetObj(UserIndex)

    End With
        
    Exit Sub

HandlePickUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePickUp", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSafeToggle_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Seguro Then
            Call WriteSafeModeOff(UserIndex)
            
        Else
            Call WriteSafeModeOn(UserIndex)

        End If
        
        .flags.Seguro = Not .flags.Seguro

    End With

    Exit Sub

HandleSafeToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

' Handles the "PartySafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePartyToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandlePartyToggle_Err
        
    '***************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    '***************************************************
    With UserList(UserIndex)
        
        .flags.SeguroParty = Not .flags.SeguroParty
        
        If .flags.SeguroParty Then
            Call WritePartySafeOn(UserIndex)
            
        Else
            Call WritePartySafeOff(UserIndex)

        End If

    End With

    Exit Sub

HandlePartyToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartyToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleSeguroClan(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSeguroClan_Err

    '***************************************************
    'Author: Ladder
    'Date: 31/10/20
    '***************************************************
    With UserList(UserIndex)

        .flags.SeguroClan = Not .flags.SeguroClan

        Call WriteClanSeguro(UserIndex, .flags.SeguroClan)

    End With

    Exit Sub

HandleSeguroClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSeguroClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleRequestGuildLeaderInfo_Err

    Call modGuilds.SendGuildLeaderInfo(UserIndex)

    Exit Sub

HandleRequestGuildLeaderInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
        
    On Error GoTo HandleRequestAtributes_Err

    Call WriteAttributes(UserIndex)

    Exit Sub

HandleRequestAtributes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestAtributes", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
        
    On Error GoTo HandleRequestSkills_Err

    Call WriteSendSkills(UserIndex)

    Exit Sub

HandleRequestSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestSkills", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
        
    On Error GoTo HandleRequestMiniStats_Err

    Call WriteMiniStats(UserIndex)

    Exit Sub

HandleRequestMiniStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestMiniStats", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleCommerceEnd_Err

    'User quits commerce mode
    If UserList(UserIndex).flags.TargetNPC <> 0 Then
    
        If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
            Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

        End If

    End If

    UserList(UserIndex).flags.Comerciando = False

    Call WriteCommerceEnd(UserIndex)
 
    Exit Sub

HandleCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUserCommerceEnd_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        
        
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user

        End If
        
        Call FinComerciarUsu(UserIndex)

    End With
        
    Exit Sub

HandleUserCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankEnd_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'User exits banking mode
        .flags.Comerciando = False
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("171", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteBankEnd(UserIndex)

    End With
        
    Exit Sub

HandleBankEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
        
    On Error GoTo HandleUserCommerceOk_Err

    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
        
    Exit Sub

HandleUserCommerceOk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOk", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUserCommerceReject_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)

        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user

            End If

        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)

    End With
        
    Exit Sub

HandleUserCommerceReject_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceReject", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Drop" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDrop_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
    '***************************************************

    Dim slot   As Byte
    Dim amount As Long
    
    With UserList(UserIndex)

        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadLong()

        If Not IntervaloPermiteTirar(UserIndex) Then Exit Sub

        If amount <= 0 Then Exit Sub

        'low rank admins can't drop item. Neither can the dead nor those sailing or riding a horse.
        If .flags.Muerto = 1 Then Exit Sub
                      
        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub
    
        'Si esta navegando y no es pirata, no dejamos tirar items al agua.
        If .flags.Navegando = 1 And Not .clase = eClass.Pirat Then
            Call WriteConsoleMsg(UserIndex, "Solo los Piratas pueden tirar items en altamar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
            
        If .flags.Montado = 1 Then
            Call WriteConsoleMsg(UserIndex, "Debes descender de tu montura para dejar objetos en el suelo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Are we dropping gold or other items??
        If slot = FLAGORO Then
            Call TirarOro(amount, UserIndex)
            
        Else
        
            '04-05-08 Ladder
            If (.flags.Privilegios And PlayerType.Admin) <> 16 Then
                If EsNewbie(UserIndex) And ObjData(.Invent.Object(slot).ObjIndex).Newbie = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No se pueden tirar los objetos Newbies.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If ObjData(.Invent.Object(slot).ObjIndex).Instransferible = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If ObjData(.Invent.Object(slot).ObjIndex).Intirable = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
            If ObjData(.Invent.Object(slot).ObjIndex).OBJType = eOBJType.otBarcos And UserList(UserIndex).flags.Navegando Then
                Call WriteConsoleMsg(UserIndex, "Para tirar la barca deberias estar en tierra firme.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            '04-05-08 Ladder
        
            'Only drop valid slots
            If slot <= UserList(UserIndex).CurrentInventorySlots And slot > 0 Then
            
                If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub

                Call DropObj(UserIndex, slot, amount, .Pos.Map, .Pos.X, .Pos.Y)

            End If

        End If

    End With
        
    Exit Sub

HandleDrop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDrop", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CastSpell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCastSpell_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim Spell As Byte
            Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Or .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
        
        If .flags.Hechizo <> 0 Then

            If (.flags.Privilegios And PlayerType.Consejero) = 0 Then

                Dim uh As Integer
                
                uh = .Stats.UserHechizos(Spell)
    
                If Hechizos(uh).AutoLanzar = 1 Then
                    UserList(UserIndex).flags.TargetUser = UserIndex
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                Else
                    Call WriteWorkRequestTarget(UserIndex, eSkill.Magia)
    
                End If
                    
            End If

        End If
        
    End With
        
    Exit Sub

HandleCastSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCastSpell", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "LeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
        
    On Error GoTo HandleLeftClick_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim X As Byte
        Dim Y As Byte
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Call LookatTile(UserIndex, .Pos.Map, X, Y)

    End With

    Exit Sub

HandleLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLeftClick", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDoubleClick_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim X As Byte
        Dim Y As Byte
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Call Accion(UserIndex, .Pos.Map, X, Y)

    End With
        
    Exit Sub

HandleDoubleClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDoubleClick", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Work" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
        
    On Error GoTo HandleWork_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    '13/01/2010: ZaMa - El pirata se puede ocultar en barca
    '***************************************************

    With UserList(UserIndex)

        Dim Skill As eSkill
            Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill

            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(UserIndex, Skill)

            Case Ocultarse

                If .flags.Montado = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás montado.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3

                    End If

                    '[/CDT]
                    Exit Sub

                End If

                If .flags.Oculto = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteLocaleMsg(UserIndex, "55", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2

                    End If

                    '[/CDT]
                    Exit Sub

                End If
                    
                If .flags.EnReto Then
                    Call WriteConsoleMsg(UserIndex, "No podés ocultarte durante un reto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                    
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
    
                End If
                    
                If MapInfo(.Pos.Map).SinInviOcul Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza divina te impide ocultarte en esta zona.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                    
                Call DoOcultarse(UserIndex)

        End Select

    End With
        
    Exit Sub

HandleWork_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWork", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUseSpellMacro_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteShowMessageBox(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        
        Call CloseSocket(UserIndex)

    End With
        
    Exit Sub

HandleUseSpellMacro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUseSpellMacro", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo HandleUseItem_Err
    
    With UserList(UserIndex)

        Dim slot As Byte
            slot = .incomingData.ReadByte()
        
        If slot <= UserList(UserIndex).CurrentInventorySlots And slot > 0 Then
            If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub

            Call UseInvItem(UserIndex, slot)

        End If

    End With

    Exit Sub

HandleUseItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUseItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCraftBlacksmith_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex).incomingData

        Dim Item As Integer
            Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item)

    End With
        
    Exit Sub

HandleCraftBlacksmith_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCraftBlacksmith", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCraftCarpenter_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex).incomingData

        Dim Item As Integer
            Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub

        Call CarpinteroConstruirItem(UserIndex, Item)

    End With
        
    Exit Sub

HandleCraftCarpenter_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCraftCarpenter", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleCraftAlquimia(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCraftAlquimia_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex).incomingData

        Dim Item As Integer
            Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub

    End With
        
    Exit Sub

HandleCraftAlquimia_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCraftAlquimia", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCraftSastre_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex).incomingData

        Dim Item As Integer
            Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub

        Call SastreConstruirItem(UserIndex, Item)

    End With

    Exit Sub

HandleCraftSastre_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCraftSastre", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
        
    On Error GoTo HandleWorkLeftClick_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim X        As Byte
        Dim Y        As Byte

        Dim Skill    As eSkill
        Dim DummyInt As Integer

        Dim tU       As Integer   'Target user
        Dim tN       As Integer   'Target NPC
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
            
        If .flags.Muerto = 1 Or .flags.Descansar Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub

        End If
            
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill

                Dim consumirMunicion As Boolean

            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                'Check Magic interval
                If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent

                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).Proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                    ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
                        DummyInt = 1

                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub

                    End If

                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                    
                    'Si no es GM invisible, le envio el movimiento del arma.
                    If UserList(UserIndex).flags.AdminInvisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))

                    End If
                    
                Else
                    Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "Estís muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                    Exit Sub

                End If
                
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                consumirMunicion = False

                'Validate target
                If tU > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub

                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "¡No podés atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub

                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    
                    Dim backup    As Byte
                    Dim envie     As Boolean
                    Dim Particula As Integer
                    Dim Tiempo    As Long

                    ' Porque no es HandleAttack ???
                    Call UsuarioAtacaUsuario(UserIndex, tU)

                    If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
                        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateFX(UserList(tU).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                    End If
                    
                    If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
                    
                        Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                        Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, Particula, Tiempo, False))

                    End If
                    
                    consumirMunicion = True
                    
                ElseIf tN > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(NpcList(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(NpcList(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub

                    End If
                    
                    'Is it attackable???
                    If NpcList(tN).Attackable <> 0 Then
                        If PuedeAtacarNPC(UserIndex, tN) Then
                            Call UsuarioAtacaNpc(UserIndex, tN)
                            consumirMunicion = True
                        Else
                            consumirMunicion = False

                        End If

                    End If

                End If
                
                With .Invent
                    DummyInt = .MunicionEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    If consumirMunicion Then
                        Call QuitarUserInvItem(UserIndex, DummyInt, 1)

                    End If
                    
                    If .Object(DummyInt).amount > 0 Then
                        'QuitarUserInvItem unequipps the ammo, so we equip it again
                        .MunicionEqpSlot = DummyInt
                        .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                        .Object(DummyInt).Equipped = 1
                    Else
                        .MunicionEqpSlot = 0
                        .MunicionEqpObjIndex = 0

                    End If

                    Call UpdateUserInv(False, UserIndex, DummyInt)

                End With

                '-----------------------------------
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                '  If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                '  Exit Sub
                ' End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub

                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check attack-spell interval
                If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Pescar
                
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 1      ' Subtipo: Caña de Pescar

                        If (MapData(.Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
                            If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).Blocked And FLAG_AGUA) <> 0 Then

                                Call DoPescar(UserIndex, False, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                            Else
                                Call WriteConsoleMsg(UserIndex, "Acércate a la costa para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If
                            
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
    
                        End If
                    
                    Case 2      ' Subtipo: Red de Pesca
    
                        If (MapData(.Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
                            
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 8 Then
                                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
    
                            End If
                                
                            If UserList(UserIndex).Stats.UserSkills(eSkill.Pescar) < 80 Then
                                Call WriteConsoleMsg(UserIndex, "Para utilizar la red de pesca debes tener 80 skills en recoleccion.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
    
                            End If
                                    
                            If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                                Call WriteConsoleMsg(UserIndex, "Esta prohibida la pesca masiva en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
    
                            End If
                                    
                            If UserList(UserIndex).flags.Navegando = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Necesitas estar sobre tu barca para utilizar la red de pesca.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
    
                            End If
                                    
                            Call DoPescar(UserIndex, True, True)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                        
                        Else
                        
                            Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
    
                        End If
                
                End Select
                    
            Case eSkill.Talar
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub

                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
        
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 6      ' Herramientas de Carpinteria - Hacha

                        ' Ahora se puede talar en la ciudad
                        'If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                        '    Call WriteConsoleMsg(UserIndex, "Esta prohibido talar arboles en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                        '    Call WriteWorkRequestTarget(UserIndex, 0)
                        '    Exit Sub
                        'End If
                            
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If
                                
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(UserIndex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If

                            If ObjData(DummyInt).Elfico <> ObjData(.Invent.HerramientaEqpObjIndex).Elfico Then
                                Call WriteConsoleMsg(UserIndex, "Sólo puedes talar árboles elficos con un hacha élfica.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If

                            If MapData(.Pos.Map, X, Y).ObjInfo.amount <= 0 Then
                                Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas leña.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Call WriteMacroTrabajoToggle(UserIndex, False)
                                Exit Sub

                            End If

                            '¡Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                Call DoTalar(UserIndex, X, Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)

                            If UserList(UserIndex).Counters.Trabajando > 1 Then
                                Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If

                        End If
                
                End Select
            
            Case eSkill.Alquimia
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 3  ' Herramientas de Alquimia - Tijeras

                        If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteConsoleMsg(UserIndex, "Esta prohibido cortar raices en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                            
                        If MapData(.Pos.Map, X, Y).ObjInfo.amount <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas raices.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                            Exit Sub

                        End If
                
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then
                            
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If
                                
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(UserIndex, "No podés quitar raices allí.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If
                                
                            '¡Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TIJERAS, .Pos.X, .Pos.Y))
                                Call DoRaices(UserIndex, X, Y)

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If
                
                End Select
                
            Case eSkill.Mineria
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 8  ' Herramientas de Mineria - Piquete
                
                        'Target whatever is in the tile
                        Call LookatTile(UserIndex, .Pos.Map, X, Y)
                            
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then

                            'Check distance
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If

                            '¡Hay un yacimiento donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then

                                ' Si el Yacimiento requiere herramienta `Dorada` y la herramienta no lo es, o vice versa.
                                ' Se usa para el yacimiento de Oro.
                                If ObjData(DummyInt).Dorada <> ObjData(.Invent.HerramientaEqpObjIndex).Dorada Then
                                    Call WriteConsoleMsg(UserIndex, "El pico dorado solo puede extraer minerales del yacimiento de Oro.", FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

                                If MapData(.Pos.Map, X, Y).ObjInfo.amount <= 0 Then
                                    Call WriteConsoleMsg(UserIndex, "Este yacimiento no tiene mas minerales para entregar.", FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub

                                End If

                                Call DoMineria(UserIndex, X, Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                            Else
                                Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                End Select

            Case eSkill.Robar

                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Seguro = 0 Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.user Then
                            If UserList(tU).flags.Muerto = 0 Then
                                Dim DistanciaMaxima As Integer

                                If .clase = eClass.Thief Then
                                    DistanciaMaxima = 2
                                Else
                                    DistanciaMaxima = 1

                                End If

                                If Abs(.Pos.X - UserList(tU).Pos.X) + Abs(.Pos.Y - UserList(tU).Pos.Y) > DistanciaMaxima Then
                                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, UserList(tU).Pos.X, UserList(tU).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                 
                                Call DoRobar(UserIndex, tU)

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "¡No podés robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)

                End If
                    
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de criaturas hostiles.
                    
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                    
                If tN > 0 Then
                    If NpcList(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 4 Then
                            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
                            
                        If LenB(NpcList(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
                            
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)
    
                End If
               
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                'Check interval
                If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                            Exit Sub

                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If
                            
                            ''FUISTE
                            Call WriteShowMessageBox(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            
                            Call CloseSocket(UserIndex)
                            Exit Sub

                        End If
                        
                        Call FundirMineral(UserIndex)
                        
                    Else
                    
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)

                        If UserList(UserIndex).Counters.Trabajando > 1 Then
                            Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If

                    End If

                Else
                
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)

                    If UserList(UserIndex).Counters.Trabajando > 1 Then
                        Call WriteMacroTrabajoToggle(UserIndex, False)

                    End If

                End If

            Case eSkill.Grupo
                'If UserList(UserIndex).Grupo.EnGrupo = False Then
                'Target whatever is in that tile
                'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser
                    
                'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
                If tU > 0 And tU <> UserIndex Then

                    'Can't steal administrative players
                    If UserList(UserIndex).Grupo.EnGrupo = False Then
                        If UserList(tU).flags.Muerto = 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 8 Then
                                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub

                            End If
                                         
                            If UserList(UserIndex).Grupo.CantidadMiembros = 0 Then
                                UserList(UserIndex).Grupo.Lider = UserIndex
                                UserList(UserIndex).Grupo.Miembros(1) = UserIndex
                                UserList(UserIndex).Grupo.CantidadMiembros = 1
                                Call InvitarMiembro(UserIndex, tU)
                            Else
                                UserList(UserIndex).Grupo.Lider = UserIndex
                                Call InvitarMiembro(UserIndex, tU)

                            End If
                                         
                        Else
                            Call WriteLocaleMsg(UserIndex, "7", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else

                        If UserList(UserIndex).Grupo.Lider = UserIndex Then
                            Call InvitarMiembro(UserIndex, tU)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Tu no podés invitar usuarios, debe hacerlo " & UserList(UserList(UserIndex).Grupo.Lider).name & ".", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    End If

                Else
                    Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                End If

                ' End If
            Case eSkill.MarcaDeClan

                'If UserList(UserIndex).Grupo.EnGrupo = False Then
                'Target whatever is in that tile
                Dim clan_nivel As Byte
                
                If UserList(UserIndex).GuildIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                
                clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

                If clan_nivel < 4 Then
                    Call WriteConsoleMsg(UserIndex, "Servidor> El nivel de tu clan debe ser 4 para utilizar esta opción.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser

                If tU = 0 Then Exit Sub
                    
                If UserList(UserIndex).GuildIndex = UserList(tU).GuildIndex Then
                    Call WriteConsoleMsg(UserIndex, "Servidor> No podes marcar a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                    
                'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
                If tU > 0 And tU <> UserIndex Then

                    ' WyroX: No puede marcar admins invisibles
                    If UserList(tU).flags.AdminInvisible <> 0 Then Exit Sub

                    'Can't steal administrative players
                    If UserList(tU).flags.Muerto = 0 Then

                        'call marcar
                        If UserList(tU).flags.invisible = 1 Or UserList(tU).flags.Oculto = 1 Then
                            Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 50, False))
                        Else
                            Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 150, False))

                        End If

                        Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageConsoleMsg("Clan> [" & UserList(UserIndex).name & "] marco a " & UserList(tU).name & ".", FontTypeNames.FONTTYPE_GUILD))
                    Else
                        Call WriteLocaleMsg(UserIndex, "7", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
                        Call WriteWorkRequestTarget(UserIndex, 0)

                    End If

                Else
                    Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eSkill.MarcaDeGM
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser

                If tU > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Servidor> [" & UserList(tU).name & "] seleccionado.", FontTypeNames.FONTTYPE_SERVER)
                Else
                    Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                End If
                    
        End Select

    End With
        
    Exit Sub

HandleWorkLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWorkLeftClick", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim Desc       As String
        Dim GuildName  As String
        Dim errorStr   As String
        Dim Alineacion As Byte
        
        Desc = .incomingData.ReadASCIIString()
        GuildName = .incomingData.ReadASCIIString()
        Alineacion = .incomingData.ReadByte()
        
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Alineacion, errorStr) Then

            Call QuitarObjetos(407, 1, UserIndex)
            Call QuitarObjetos(408, 1, UserIndex)
            Call QuitarObjetos(409, 1, UserIndex)
            Call QuitarObjetos(411, 1, UserIndex)

126             Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.name & " fundó el clan " & GuildName & " con alineación " & IIf(Alineacion = 0, "ciudadana", "criminal") & ".", FontTypeNames.FONTTYPE_GUILD))
128             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
                'Update tag
130             Call RefreshCharStatus(UserIndex)
            Else
132             Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateNewGuild", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "SpellInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSpellInfo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim spellSlot As Byte
        Dim Spell     As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)

        If Spell > 0 And Spell < NumeroHechizos + 1 Then

            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "HECINF*" & Spell, FontTypeNames.FONTTYPE_INFO)

            End With

        End If

    End With
        
    Exit Sub

HandleSpellInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpellInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "EquipItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
        
    On Error GoTo HandleEquipItem_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim itemSlot As Byte
            itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate item slot
        If itemSlot > UserList(UserIndex).CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)

    End With
        
    Exit Sub

HandleEquipItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEquipItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeHeading_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/01/08
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Heading As eHeading
            Heading = .incomingData.ReadByte()
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End If

    End With

    Exit Sub

HandleChangeHeading_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeHeading", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
        
    On Error GoTo HandleModifySkills_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim i                      As Long
        Dim Count                  As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub

            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats

            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100

                End If

            Next i

        End With

    End With
        
    Exit Sub

HandleModifySkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleModifySkills", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Train" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
        
    On Error GoTo HandleTrain_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim SpawnedNpc As Integer
        Dim PetIndex   As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If NpcList(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
        
            If PetIndex > 0 And PetIndex < NpcList(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(NpcList(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, NpcList(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    NpcList(SpawnedNpc).MaestroNPC = .flags.TargetNPC
                    NpcList(.flags.TargetNPC).Mascotas = NpcList(.flags.TargetNPC).Mascotas + 1

                End If

            End If

        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))

        End If

    End With
        
    Exit Sub

HandleTrain_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrain", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCommerceBuy_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim slot   As Byte
        Dim amount As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub

        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Call WriteCommerceEnd(UserIndex)
            Exit Sub

        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, slot, amount)

    End With

    Exit Sub

HandleCommerceBuy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceBuy", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankExtractItem_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        slotdestino = .incomingData.ReadByte()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

        'User retira el item del slot
        Call UserRetiraItem(UserIndex, slot, amount, slotdestino)

    End With

    Exit Sub

HandleBankExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankExtractItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCommerceSell_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim slot   As Byte
        Dim amount As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'íEl target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub

        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, slot, amount)

    End With

    Exit Sub

HandleCommerceSell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceSell", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankDeposit_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        slotdestino = .incomingData.ReadByte()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'íEl target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
            
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, slot, amount, slotdestino)

    End With
        
    Exit Sub

HandleBankDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankDeposit", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ForumPost" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim File     As String
        Dim title    As String
        Dim msg      As String
        Dim postFile As String
        Dim handle   As Integer
        Dim i        As Long
        Dim Count    As Integer
        
        title = .incomingData.ReadASCIIString()
        msg = .incomingData.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            File = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
            If FileExist(File, vbNormal) Then
                Count = val(GetVar(File, "INFO", "CantMSG"))
                
                'If there are too many messages, delete the forum
                If Count > MAX_MENSAJES_FORO Then

                    For i = 1 To Count
                        Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
                    Next i

                    Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
                    Count = 0

                End If

            Else
                'Starting the forum....
                Count = 0

            End If
            
            handle = FreeFile()
            postFile = Left$(File, Len(File) - 4) & CStr(Count + 1) & ".for"
            
            'Create file
            Open postFile For Output As handle
            Print #handle, title
            Print #handle, msg
            Close #handle
            
            'Update post count
            Call WriteVar(File, "INFO", "CantMSG", Count + 1)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Close #handle
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForumPost", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
        
    On Error GoTo HandleMoveSpell_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex).incomingData

        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1

        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())

    End With
        
    Exit Sub

HandleMoveSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Desc As String
        
        Desc = .incomingData.ReadASCIIString()
        
        Call modGuilds.ChangeCodexAndDesc(Desc, .GuildIndex)

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUserCommerceOffer_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim tUser  As Integer
        Dim slot   As Byte
        Dim amount As Long
            
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadLong()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If Amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((slot < 1 Or slot > UserList(UserIndex).CurrentInventorySlots) And slot <> FLAGORO) Or amount <= 0 Then Exit Sub
        
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
        'Is the commerce attempt valid??
        If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub

        End If
        
        'Is he still logged??
        If Not UserList(tUser).flags.UserLogged Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else

            'Is he alive??
            If UserList(tUser).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub

            End If
            
            'Has he got enough??
            If slot = FLAGORO Then

                'gold
                If amount > .Stats.GLD Then
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            Else

                'inventory
                If amount > .Invent.Object(slot).amount Then
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
            
            'Prevent offer changes (otherwise people would ripp off other players)
            'If .ComUsu.Objeto > 0 Then
            '     Call WriteConsoleMsg(UserIndex, "No podés cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
            '     Exit Sub

            '  End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = slot Then
                    Call WriteConsoleMsg(UserIndex, "No podés vender tu barco mientras lo estás usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
            
            If .flags.Montado = 1 Then
                If .Invent.MonturaSlot = slot Then
                    Call WriteConsoleMsg(UserIndex, "No podés vender tu montura mientras la estás usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
                
            If .Invent.Object(slot).ObjIndex > 0 Then
                If ObjData(.Invent.Object(slot).ObjIndex).Instransferible Then
                    Call WriteConsoleMsg(UserIndex, "Este objeto es intransferible, no podés venderlo.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
            
            .ComUsu.Objeto = slot
            .ComUsu.cant = amount
            
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto = True Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(tUser, .name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)

            End If
            
            Dim ObjAEnviar As obj
                
            ObjAEnviar.amount = amount

            'Si no es oro tmb le agrego el objInex
            If slot <> 200 Then ObjAEnviar.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
            'Llamos a la funcion
            Call EnviarObjetoTransaccion(tUser, UserIndex, ObjAEnviar)

        End If

    End With

    Exit Sub

HandleUserCommerceOffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOffer", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = .incomingData.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If

    End With
        
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptPeace", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = .incomingData.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildRejectAlliance", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = .incomingData.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildRejectPeace", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = .incomingData.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptAlliance", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = .incomingData.ReadASCIIString()
        proposal = .incomingData.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)

        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = .incomingData.ReadASCIIString()
        proposal = .incomingData.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)

        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        
        guild = .incomingData.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        
        guild = .incomingData.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
            
    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildPeaceDetails", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim user    As String
        Dim details As String
        
        user = .incomingData.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, user)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildRequestJoinerInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo HandleGuildAlliancePropList_Err
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
        
    Exit Sub

HandleGuildAlliancePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildAlliancePropList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
      
    On Error GoTo HandleGuildPeacePropList_Err

    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
        
    Exit Sub

HandleGuildPeacePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild           As String
        Dim errorStr        As String
        Dim otherGuildIndex As Integer
        
        guild = .incomingData.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo ErrHandler

    Call modGuilds.ActualizarWebSite(UserIndex, UserList(UserIndex).incomingData.ReadASCIIString())

    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildNewWebsite", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim errorStr As String
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("[" & UserName & "] ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim errorStr As String
        Dim UserName As String
        Dim Reason   As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName   As String
        Dim GuildIndex As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        Else
            Call WriteConsoleMsg(UserIndex, "No podés expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildKickMember", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    Call modGuilds.ActualizarNoticias(UserIndex, UserList(UserIndex).incomingData.ReadASCIIString())

    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildUpdateNews", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    Call modGuilds.SendDetallesPersonaje(UserIndex, UserList(UserIndex).incomingData.ReadASCIIString())

    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildMemberInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGuildOpenElections_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim Error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .name, FontTypeNames.FONTTYPE_GUILD))

        End If

    End With
        
    Exit Sub

HandleGuildOpenElections_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildOpenElections", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild       As String
        Dim application As String
        Dim errorStr    As String
        
        guild = .incomingData.ReadASCIIString()
        application = .incomingData.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildRequestMembership", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler
 
    Call modGuilds.SendGuildDetails(UserIndex, UserList(UserIndex).incomingData.ReadASCIIString())

    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildRequestDetails", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Online" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOnline_Err

    '***************************************************
    'Ladder 17/12/20 : Envio records de usuarios y uptime
    '***************************************************
        
    Dim i         As Long
    Dim Count     As Long
    Dim Time      As Long
    Dim UpTimeStr As String
    
    With UserList(UserIndex)

        Dim nombres As String
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
            
                If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
                    nombres = nombres & " - " & UserList(i).name
                End If

                Count = Count + 1

            End If

        Next i
        
        'Get total time in seconds
        Time = ((GetTickCount()) - tInicioServer) \ 1000
        
        'Get times in dd:hh:mm:ss format
        UpTimeStr = (Time Mod 60) & " segundos."
        Time = Time \ 60
        
        UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
        Time = Time \ 60
        
        UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
        Time = Time \ 24
        
        If Time = 1 Then
            UpTimeStr = Time & " día, " & UpTimeStr
        Else
            UpTimeStr = Time & " días, " & UpTimeStr
    
        End If
    
        Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)

        If .flags.Privilegios And PlayerType.user Then
            Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados.", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteConsoleMsg(UserIndex, "Tiempo en línea: " & UpTimeStr & " Record de usuarios en simultaneo: " & RecordUsuarios & ".", FontTypeNames.FONTTYPE_INFOIAO)

        Else
            Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados: " & nombres & ".", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteConsoleMsg(UserIndex, "Tiempo en línea: " & UpTimeStr & " Record de usuarios en simultaneo: " & RecordUsuarios & ".", FontTypeNames.FONTTYPE_INFOIAO)

        End If

    End With
        
    Exit Sub

HandleOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOnline", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Quit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
        
    On Error GoTo HandleQuit_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '04/15/2008 - No se reseteaban los contadores de invi ni de ocultar. (NicoNZ)
    '***************************************************
        
    Dim tUser        As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)

        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podés salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
            
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)

                End If

            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)

        End If
        
        isNotVisible = (.flags.Oculto Or .flags.invisible)

        If isNotVisible And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .flags.invisible = 0

            .Counters.Invisibilidad = 0
            .Counters.TiempoOculto = 0
                
            'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

        End If

        Call Cerrar_Usuario(UserIndex)

    End With

    Exit Sub

HandleQuit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuit", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGuildLeave_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)

        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no podés salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

    Exit Sub

HandleGuildLeave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildLeave", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRequestAccountState_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim earnings   As Integer
    Dim percentage As Integer
    
    With UserList(UserIndex)

        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Select Case NpcList(.flags.TargetNPC).NPCtype

            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tenes " & PonerPuntos(.Stats.Banco) & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero

                If Not .flags.Privilegios And PlayerType.user Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)

                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)

                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & PonerPuntos(Apuestas.Ganancias) & " Salida: " & PonerPuntos(Apuestas.Perdidas) & " Ganancia Neta: " & PonerPuntos(earnings) & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)

                End If

        End Select

    End With
        
    Exit Sub

HandleRequestAccountState_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestAccountState", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo HandlePetStand_Err
        
    With UserList(UserIndex)

        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's his pet
        If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        NpcList(.flags.TargetNPC).Movement = TipoAI.Estatico
        
        Call Expresar(.flags.TargetNPC, UserIndex)

    End With
        
    Exit Sub

HandlePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePetStand", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo HandlePetFollow_Err
        
    With UserList(UserIndex)

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make usre it's the user's pet
        If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)

    End With
        
    Exit Sub

HandlePetFollow_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePetFollow", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "PetLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetLeave(ByVal UserIndex As Integer)
    '***************************************************
        
    On Error GoTo HandlePetLeave_Err
        
    With UserList(UserIndex)

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make usre it's the user's pet
        If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub

        Call QuitarNPC(.flags.TargetNPC)

    End With
        
    Exit Sub

HandlePetLeave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePetLeave", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GrupoMsg" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGrupoMsg(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String
            chat = .incomingData.ReadASCIIString()
        
        If LenB(chat) <> 0 Then

            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .Grupo.EnGrupo = True Then

                Dim i As Byte
         
                For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
                    
                    'Call WriteConsoleMsg(UserList(.Grupo.Lider).Grupo.Miembros(i), "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
                    Call WriteConsoleMsg(UserList(.Grupo.Lider).Grupo.Miembros(i), .name & "> " & chat, FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
                    Call WriteChatOverHead(UserList(.Grupo.Lider).Grupo.Miembros(i), chat, UserList(UserIndex).Char.CharIndex, &HFF8000)
                  
                Next i
            
            Else
                'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_New_GRUPO)
                Call WriteConsoleMsg(UserIndex, "Grupo> No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGrupoMsg", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "TrainList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
        
    On Error GoTo HandleTrainList_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's the trainer
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)

    End With

    Exit Sub

HandleTrainList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrainList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Rest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRest_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzís a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else

            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub

            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    Exit Sub

HandleRest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Meditate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
        
    On Error GoTo HandleMeditate_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arreglí un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
        
    With UserList(UserIndex)

        'Si ya tiene el mana completo, no lo dejamos meditar.
        If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
                           
        'Las clases NO MAGICAS no meditan...
        If .clase = eClass.Hunter Or .clase = eClass.Trabajador Or .clase = eClass.Warrior Or .clase = eClass.Pirat Or .clase = eClass.Thief Then Exit Sub

        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.Montado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podes meditar estando montado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        .flags.Meditando = Not .flags.Meditando

        If .flags.Meditando Then

            .Counters.TimerMeditar = 0

            Select Case .Stats.ELV

                Case 1 To 14
                    .Char.FX = Meditaciones.MeditarInicial

                Case 15 To 29
                    .Char.FX = Meditaciones.MeditarMayor15

                Case 30 To 39
                    .Char.FX = Meditaciones.MeditarMayor30

                Case 40 To 44
                    .Char.FX = Meditaciones.MeditarMayor40

                Case 45 To 46
                    .Char.FX = Meditaciones.MeditarMayor45

                Case Else
                    .Char.FX = Meditaciones.MeditarMayor47

            End Select

        Else
            .Char.FX = 0

            'Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, .Char.FX))

    End With
        
    Exit Sub

HandleMeditate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMeditate", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Resucitate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
        
    On Error GoTo HandleResucitate_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate NPC and make sure player is dead
        If (NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (NpcList(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call RevivirUsuario(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleResucitate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleResucitate", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Heal" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
        
    On Error GoTo HandleHeal_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If (NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And NpcList(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "ííHas sido curado!!", FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleHeal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHeal", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleRequestStats_Err

    Call SendUserStatsTxt(UserIndex, UserIndex)
        
    Exit Sub

HandleRequestStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestStats", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Help" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleHelp_Err

    Call SendHelp(UserIndex)
        
    Exit Sub

HandleHelp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHelp", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCommerceStart_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
                
            'VOS, como GM, NO podes COMERCIAR con NPCs. (excepto Dioses y Admins)
            If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                Call WriteConsoleMsg(UserIndex, "No podés vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
                
            'Does the NPC want to trade??
            If NpcList(.flags.TargetNPC).Comercia = 0 Then
                If LenB(NpcList(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                End If
                
                Exit Sub

            End If
            
            If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
                
        ElseIf .flags.TargetUser > 0 Then

            ' **********************  Comercio con Usuarios  *********************
                
            'VOS, como GM, NO podes COMERCIAR con usuarios. (excepto Dioses y Admins)
            If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                Call WriteConsoleMsg(UserIndex, "No podés vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
                
            'NO podes COMERCIAR CON un GM. (excepto Dioses y Admins)
            If (UserList(.flags.TargetUser).flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                Call WriteConsoleMsg(UserIndex, "No podés vender items a este usuario.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
                
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No podés comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No podés comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No podés comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name
            .ComUsu.cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleCommerceStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceStart", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankStart_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 6 Then
                Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'If it's the banker....
            If NpcList(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleBankStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankStart", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Enlist" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
        
    On Error GoTo HandleEnlist_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte mís.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
            
        Else
            Call EnlistarCaos(UserIndex)

        End If

    End With
        
    Exit Sub

HandleEnlist_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEnlist", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Information" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
        
    On Error GoTo HandleInformation_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legiín oscura!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

    End With
        
    Exit Sub

HandleInformation_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInformation", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Reward" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
        
    On Error GoTo HandleReward_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
        
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call RecompensaArmadaReal(UserIndex)
            
        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legiín oscura!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call RecompensaCaos(UserIndex)

        End If

    End With
        
    Exit Sub

HandleReward_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleReward", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleRequestMOTD_Err

    Call SendMOTD(UserIndex)
        
    Exit Sub

HandleRequestMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestMOTD", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "UpTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/08
    '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
    '***************************************************
   
    On Error GoTo HandleUpTime_Err

    Dim Time      As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    Time = ((GetTickCount()) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr

    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
        
    Exit Sub

HandleUpTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpTime", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Inquiry" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleInquiry_Err

    Call ConsultaPopular.SendInfoEncuesta(UserIndex)
        
    Exit Sub

HandleInquiry_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInquiry", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String
            chat = .incomingData.ReadASCIIString()
        
        If LenB(chat) <> 0 Then

            'Analize chat...
            Call Statistics.ParseChat(chat)

            ' WyroX: Foto-denuncias - Push message
            Dim i As Integer

            For i = 1 To UBound(.flags.ChatHistory) - 1
                .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
            Next
                
            .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat))

                'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                'Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "CentinelReport" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCentinelReport_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call CentinelaCheckClave(UserIndex, UserList(UserIndex).incomingData.ReadInteger())

    Exit Sub

HandleCentinelReport_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCentinelReport", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGuildOnline_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim onlineList As String
            onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Compaíeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)

        End If

    End With
        
    Exit Sub

HandleGuildOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildOnline", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String
            chat = .incomingData.ReadASCIIString()
        
        If LenB(chat) <> 0 Then

            'Analize chat...
            Call Statistics.ParseChat(chat)

            ' WyroX: Foto-denuncias - Push message
            Dim i As Long
            For i = 1 To UBound(.flags.ChatHistory) - 1
                .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
            Next
                
            .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))

            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCouncilMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim request As String
            request = .incomingData.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If

    End With
    
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRoleMasterRequest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GMRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGMRequest_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    With UserList(UserIndex)

        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sílo debes esperar que se desocupe algín GM.", FontTypeNames.FONTTYPE_INFO)
                
        Else
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleGMRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGMRequest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Description As String
            Description = .incomingData.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFOIAO)

        Else
            
            If Len(Description) > 128 Then
                Call WriteConsoleMsg(UserIndex, "La descripción es muy larga.", FontTypeNames.FONTTYPE_INFOIAO)

            ElseIf Not DescripcionValida(Description) Then
                Call WriteConsoleMsg(UserIndex, "La descripción tiene carácteres inválidos.", FontTypeNames.FONTTYPE_INFOIAO)
                
            Else
                .Desc = Trim$(Description)
                Call WriteConsoleMsg(UserIndex, "La descripción a cambiado.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeDescription", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GuildVote" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim vote     As String
        Dim errorStr As String
        
        vote = .incomingData.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)

        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
 
    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildVote", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Punishments" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim name As String
            name = .incomingData.ReadASCIIString()

        ' Si un GM usa este comando, me fijo que me haya dado el nick del PJ a analizar.
        If EsGM(UserIndex) And LenB(name) = 0 Then Exit Sub
        
        Dim Count As Integer

        If (InStrB(name, "\") <> 0) Then
            name = Replace(name, "\", vbNullString)

        End If

        If (InStrB(name, "/") <> 0) Then
            name = Replace(name, "/", vbNullString)

        End If

        If (InStrB(name, ":") <> 0) Then
            name = Replace(name, ":", vbNullString)

        End If

        If (InStrB(name, "|") <> 0) Then
            name = Replace(name, "|", vbNullString)

        End If
           
        Dim TargetUserName As String

        If EsGM(UserIndex) Then
        
            If PersonajeExiste(name) Then
                TargetUserName = name
                
            Else
                Call WriteConsoleMsg(UserIndex, "El personaje " & TargetUserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
        Else
        
            TargetUserName = .name
            
        End If

        If Database_Enabled Then
            Count = GetUserAmountOfPunishmentsDatabase(TargetUserName)
                
        Else
            Count = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))

        End If

        If Count = 0 Then
            Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)

        Else
                
            If Database_Enabled Then
                Call SendUserPunishmentsDatabase(UserIndex, TargetUserName)
                        
            Else
                        
                While Count > 0

                    Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & TargetUserName & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                    Count = Count - 1
                Wend
                            
            End If

        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePunishments", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ChangePassword" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Creation Date: 10/10/07
    'Last Modified By: Ladder
    'Ahora cambia la password de la cuenta y no del PJ.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim oldPass  As String
        Dim newPass  As String
        Dim oldPass2 As String

        oldPass = .incomingData.ReadASCIIString()
        newPass = .incomingData.ReadASCIIString()

        Call ChangePasswordDatabase(UserIndex, SDesencriptar(oldPass), SDesencriptar(newPass))

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangePassword", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Gamble" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGamble_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim amount As Integer
            amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

        ElseIf Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                
        ElseIf NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        ElseIf amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        ElseIf amount > 10000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 10.000 monedas.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        ElseIf .Stats.GLD < amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        Else

            If RandomNumber(1, 100) <= 45 Then
                .Stats.GLD = .Stats.GLD + amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & PonerPuntos(amount) & " monedas de oro!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & PonerPuntos(amount) & " monedas de oro.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)

        End If

    End With

    Exit Sub

HandleGamble_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGamble", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
        
    On Error GoTo HandleInquiryVote_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim opt As Byte
            opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)

    End With
        
    Exit Sub

HandleInquiryVote_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInquiryVote", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankExtractGold_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim amount As Long
            amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If amount > 0 And amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - amount
            .Stats.GLD = .Stats.GLD + amount
            'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            Call WriteUpdateGold(UserIndex)
            Call WriteGoliathInit(UserIndex)

        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

    End With

    Exit Sub

HandleBankExtractGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankExtractGold", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
        
    On Error GoTo HandleLeaveFaction_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    With UserList(UserIndex)

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .Faccion.ArmadaReal = 0 And .Faccion.FuerzasCaos = 0 Then
            If .Faccion.Status = 1 Then
                Call VolverCriminal(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Ahora sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

        Else

            ' Call WriteConsoleMsg(UserIndex, "Ya sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
            ' Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then

            If .Faccion.ArmadaReal = 1 Then
                Call WriteConsoleMsg(UserIndex, "Para salir del ejercito debes ir a visitar al rey.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            ElseIf .Faccion.FuerzasCaos = 1 Then
                Call WriteConsoleMsg(UserIndex, "Para salir de la legion debes ir a visitar al diablo.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

            Exit Sub

        End If
        
        If NpcList(.flags.TargetNPC).NPCtype = eNPCType.Enlistador Then

            'Quit the Royal Army?
            If .Faccion.ArmadaReal = 1 Then
                If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
                    Call ExpulsarFaccionReal(UserIndex)
                    Call WriteChatOverHead(UserIndex, "Serís bienvenido a las fuerzas imperiales si deseas regresar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub
                Else
                    Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                   
                End If

                'Quit the Chaos Legion??
            ElseIf .Faccion.FuerzasCaos = 1 Then

                If NpcList(.flags.TargetNPC).flags.Faccion = 1 Then
                    Call ExpulsarFaccionCaos(UserIndex)
                    Call WriteChatOverHead(UserIndex, "Ya volverís arrastrandote.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Else
                    Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                End If

            Else
                Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End If
    
    End With
        
    Exit Sub

HandleLeaveFaction_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLeaveFaction", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBankDepositGold_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim amount As Long
            amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
            
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If amount > 0 And amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + amount
            .Stats.GLD = .Stats.GLD - amount
            'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteGoliathInit(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

    End With
        
    Exit Sub

HandleBankDepositGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankDepositGold", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Denounce" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleFinEvento(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDenounce_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then Exit Sub

        If EventoActivo Then
            Call FinalizarEvento
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ningun evento activo.", FontTypeNames.FONTTYPE_New_Eventos)
        
        End If
        
    End With
        
    Exit Sub

HandleDenounce_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub ''

' Handles the "GuildMemberList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild       As String
        Dim memberCount As Integer
        Dim i           As Long
        Dim UserName    As String
        
        guild = .incomingData.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")

            End If

            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")

            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)

            Else
                
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If
            
    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildMemberList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GMMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If

    End With
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGMMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ShowName" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
        
    On Error GoTo HandleShowName_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
        
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)

        End If

    End With
        
    Exit Sub

HandleShowName_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowName", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOnlineRoyalArmy_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
    
        Dim i    As Long
        Dim list As String

        For i = 1 To LastUser

            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Or .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "

                    End If

                End If

            End If

        Next i

    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)

    End If
        
    Exit Sub

HandleOnlineRoyalArmy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOnlineRoyalArmy", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOnlineChaosLegion_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
    
        Dim i    As Long
        Dim list As String

        For i = 1 To LastUser

            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Or .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "

                    End If

                End If

            End If

        Next i

    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)

    End If
        
    Exit Sub

HandleOnlineChaosLegion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOnlineChaosLegion", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GoNearby" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
            UserName = .incomingData.ReadASCIIString()
        
        Dim tIndex As Integer

        Dim X      As Long
        Dim Y      As Long

        Dim i      As Long
            
        Dim Found  As Boolean
        
        'Check the user has enough powers
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Or Ayuda.Existe(UserName) Then
            tIndex = NameIndex(UserName)

            'Si es dios o Admins no podemos salvo que nosotros tambiín lo seamos
            If CompararPrivilegiosUser(UserIndex, tIndex) >= 0 Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Found = True
                                        Exit For

                                    End If

                                End If

                            Next Y
                            
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares estín ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGoNearby", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Comment" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim comment As String
            comment = .incomingData.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleComment", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ServerTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
        
    On Error GoTo HandleServerTime_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
    
        Call LogGM(.name, "Hora.")

    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
        
    Exit Sub

HandleServerTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleServerTime", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Where" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And Not (PlayerType.Consejero Or PlayerType.user)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If CompararPrivilegiosUser(UserIndex, tUser) >= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)

                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWhere", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCreaturesInMap_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizaciín.
    '***************************************************

    With UserList(UserIndex)

        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1()    As String
        Dim List2()    As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.user Then Exit Sub
        
        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If NpcList(i).Pos.Map = Map Then

                    'íesta vivo?
                    If NpcList(i).flags.NPCActive And NpcList(i).Hostile = 1 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List1(j) = List1(j) & ", (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List2(j) = List2(j) & ", (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.Y & ")"
                                NPCcant2(j) = 1

                            End If

                        End If

                    End If

                End If

            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay mís NPCS", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call LogGM(.name, "Numero enemigos en mapa " & Map)

        End If

    End With
        
    Exit Sub

HandleCreaturesInMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreaturesInMap", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
        
    On Error GoTo HandleWarpMeToTarget_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
        
        Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
        
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

    End With
        
    Exit Sub

HandleWarpMeToTarget_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWarpMeToTarget", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "WarpChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Map      As Integer
        Dim X        As Byte
        Dim Y        As Byte
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        Map = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        If .flags.Privilegios And PlayerType.user Then Exit Sub
            
        If .flags.Privilegios And PlayerType.Consejero Then
        
            If MapInfo(Map).Seguro = 0 Then
                Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ciudades.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                'Si manda yo o su propio nombre
            ElseIf LCase$(UserName) <> LCase$(UserList(UserIndex).name) And UCase$(UserName) <> "YO" Then
                Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
            
        '¿Para que te vas a transportar a la misma posicion?
        If .Pos.Map = Map And .Pos.X = X And .Pos.Y = Y Then Exit Sub
            
        If MapaValido(Map) And LenB(UserName) <> 0 Then

            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
                
            Else
                tUser = UserIndex

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)

            ElseIf InMapBounds(Map, X, Y) Then
                Call FindLegalPos(tUser, Map, X, Y)
                Call WarpUserChar(tUser, Map, X, Y, True)

                If tUser <> UserIndex Then
                    Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                End If
                        
            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWarpChar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Silence" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim minutos  As Integer
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        minutos = .incomingData.ReadInteger()

        If EsGM(UserIndex) Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then

                If PersonajeExiste(UserName) Then

                    If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(UserName)) > 0 Then

                        If minutos > 0 Then
                            Call SilenciarUserDatabase(UserName, minutos)
                            Call SavePenaDatabase(UserName, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
                            Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .name & " ha silenciado a " & UserName & "(offline) por " & minutos & " minutos.", FontTypeNames.FONTTYPE_GM))
                            Call LogGM(.name, "Silenciar a " & UserList(tUser).name & " por " & minutos & " minutos.")
                        Else
                            Call DesilenciarUserDatabase(UserName)
                            Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .name & " ha desilenciado a " & UserName & "(offline).", FontTypeNames.FONTTYPE_GM))
                            Call LogGM(.name, "Desilenciar a " & UserList(tUser).name & ".")

                        End If
                            
                    Else
                        
                        Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    
                    Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            ElseIf CompararPrivilegiosUser(UserIndex, tUser) > 0 Then

                If minutos > 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    UserList(tUser).flags.MinutosRestantes = minutos
                    UserList(tUser).flags.SegundosPasados = 0

                    Call SavePenaDatabase(UserName, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .name & " ha silenciado a " & UserList(tUser).name & " por " & minutos & " minutos.", FontTypeNames.FONTTYPE_GM))
                    Call WriteConsoleMsg(tUser, "Has sido silenciado por los administradores, no podrás hablar con otros usuarios. Utilice /GM para pedir ayuda.", FontTypeNames.FONTTYPE_GM)
                    Call LogGM(.name, "Silenciar a " & UserList(tUser).name & " por " & minutos & " minutos.")

                Else
                    
                    UserList(tUser).flags.Silenciado = 1

                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .name & " ha desilenciado a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_GM))
                    Call WriteConsoleMsg(tUser, "Has sido desilenciado.", FontTypeNames.FONTTYPE_GM)
                    Call LogGM(.name, "Desilenciar a " & UserList(tUser).name & ".")

                End If
                    
            Else
                
                Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSilence", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSOSShowList_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub

        Call WriteShowSOSForm(UserIndex)

    End With
        
    Exit Sub

HandleSOSShowList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSOSShowList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
            UserName = .incomingData.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then Call Ayuda.Quitar(UserName)

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSOSRemove", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "GoToChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        Dim X        As Byte
        Dim Y        As Byte
        
        UserName = .incomingData.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            If LenB(UserName) <> 0 Then
                tUser = NameIndex(UserName)
                    
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            Else
                tUser = .flags.TargetUser

                If tUser <= 0 Then Exit Sub

            End If
      
            If CompararPrivilegiosUser(tUser, UserIndex) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser).name & " que quieres ir a su posición.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " quiere transportarse a tu ubicación. Escribe /sum " & .name & " para traerlo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            X = UserList(tUser).Pos.X
            Y = UserList(tUser).Pos.Y + 1

            Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                
            Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
            If .flags.AdminInvisible = 0 Then
                Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)

            End If
                
            Call WriteConsoleMsg(UserIndex, "Te has transportado hacia " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                    
            Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGoToChar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleDesbuggear(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String, tUser As Integer, i As Long, Count As Long
        
        UserName = .incomingData.ReadASCIIString()
        
        If EsGM(UserIndex) Then
            If Len(UserName) > 0 Then
                tUser = NameIndex(UserName)
                
                If tUser > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario debe estar offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    Dim AccountId As Long, AccountOnline As Boolean
                    
                    AccountId = GetAccountIDDatabase(UserName)
                    
                    If AccountId >= 0 Then

                        For i = 1 To LastUser

                            If UserList(i).flags.UserLogged Then
                                If UserList(i).AccountId = AccountId Then
                                    AccountOnline = True

                                End If

                                Count = Count + 1

                            End If

                        Next i
                        
                        NumUsers = Count
                        Call MostrarNumUsers
                        
                        If AccountOnline Then
                            Call WriteConsoleMsg(UserIndex, "Hay un usuario de la cuenta conectado. Se actualizaron solo los usuarios online.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ResetLoggedDatabase(AccountId)
                            Call WriteConsoleMsg(UserIndex, "Cuenta del personaje desbuggeada y usuarios online actualizados.", FontTypeNames.FONTTYPE_INFO)

                        End If
    
                        Call LogGM(.name, "/DESBUGGEAR " & UserName)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else

                For i = 1 To LastUser

                    If UserList(i).flags.UserLogged Then
                        Count = Count + 1

                    End If

                Next i
                
                NumUsers = Count
                Call MostrarNumUsers
                
                Call WriteConsoleMsg(UserIndex, "Se actualizaron los usuarios online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDesbuggear", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleDarLlaveAUsuario(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String, tUser As Integer, Llave As Integer
        
        UserName = .incomingData.ReadASCIIString()
        Llave = .incomingData.ReadInteger()
        
        ' Solo dios o admin
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then

            ' Me aseguro que esté activada la db
            If Not Database_Enabled Then
                Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)
            
                ' Me aseguro que el objeto sea una llave válida
            ElseIf Llave < 1 Or Llave > NumObjDatas Then
                Call WriteConsoleMsg(UserIndex, "El número ingresado no es el de una llave válida.", FontTypeNames.FONTTYPE_INFO)

            ElseIf ObjData(Llave).OBJType <> eOBJType.otLlaves Then ' vb6 no tiene short-circuit evaluation :(
                Call WriteConsoleMsg(UserIndex, "El número ingresado no es el de una llave válida.", FontTypeNames.FONTTYPE_INFO)

            Else
                
                tUser = NameIndex(UserName)
                
                If tUser > 0 Then

                    ' Es un user online, guardamos la llave en la db
                    If DarLlaveAUsuarioDatabase(UserName, Llave) Then

                        ' Actualizamos su llavero
                        If MeterLlaveEnLLavero(tUser, Llave) Then
                            Call WriteConsoleMsg(UserIndex, "Llave número " & Llave & " entregada a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. El usuario no tiene más espacio en su llavero.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible.", FontTypeNames.FONTTYPE_INFO)

                    End If
                        
                Else
                    
                    ' No es un usuario online, nos fijamos si es un email
                    If CheckMailString(UserName) Then

                        ' Es un email, intentamos guardarlo en la db
                        If DarLlaveACuentaDatabase(UserName, Llave) Then
                            Call WriteConsoleMsg(UserIndex, "Llave número " & Llave & " entregada a " & LCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible y que el email sea correcto.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no está online. Ingrese el email de la cuenta para otorgar la llave offline.", FontTypeNames.FONTTYPE_INFO)

                    End If
    
                End If
                
                Call LogGM(.name, "/DARLLAVE " & UserName & " " & Llave)

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDarLlaveAUsuario", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleSacarLlave(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSacarLlave_Err

    With UserList(UserIndex)

        Dim Llave As Integer
            Llave = .incomingData.ReadInteger()
        
        ' Solo dios o admin
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then

            ' Me aseguro que esté activada la db
            If Not Database_Enabled Then
                Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)

            Else

                ' Intento borrarla de la db
                If SacarLlaveDatabase(Llave) Then
                    Call WriteConsoleMsg(UserIndex, "La llave " & Llave & " fue removida.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se pudo sacar la llave. Asegúrese de que esté en uso.", FontTypeNames.FONTTYPE_INFO)

                End If

                Call LogGM(.name, "/SACARLLAVE " & Llave)

            End If
                
        End If

    End With
        
    Exit Sub

HandleSacarLlave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSacarLlave", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleVerLlaves(ByVal UserIndex As Integer)
        
    On Error GoTo HandleVerLlaves_Err

    With UserList(UserIndex)

        ' Sólo GMs
        If Not (.flags.Privilegios And PlayerType.user) Then

            ' Me aseguro que esté activada la db
            If Not Database_Enabled Then
                Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            ' Leo y muestro todas las llaves usadas
            Call VerLlavesDatabase(UserIndex)

        End If
                
    End With

    Exit Sub

HandleVerLlaves_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleVerLlaves", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleUseKey(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUseKey_Err

    With UserList(UserIndex)

        Dim slot As Byte
            slot = .incomingData.ReadByte

        Call UsarLlave(UserIndex, slot)
                
    End With
        
    Exit Sub

HandleUseKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUseKey", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Invisible" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
        
    On Error GoTo HandleInvisible_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)

    End With
        
    Exit Sub

HandleInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInvisible", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GMPanel" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGMPanel_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.user Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)

    End With
        
    Exit Sub

HandleGMPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGMPanel", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GMPanel" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRequestUserList_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i       As Long
    Dim names() As String
    Dim Count   As Long
    
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster)) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser

            If (LenB(UserList(i).name) <> 0) Then
                
                names(Count) = UserList(i).name
                Count = Count + 1
 
            End If

        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

    End With
        
    Exit Sub

HandleRequestUserList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestUserList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Working" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
        
    On Error GoTo HandleWorking_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long
    Dim Users As String
    
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster)) Then Exit Sub
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then Users = Users & " (*)"

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleWorking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWorking", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Hiding" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
        
    On Error GoTo HandleHiding_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long
    Dim Users As String
    
    With UserList(UserIndex)

        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster)) Then Exit Sub
        
        For i = 1 To LastUser

            If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).name & ", "

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleHiding_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Jail" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        
        
        
        Dim UserName As String
        Dim Reason   As String
        Dim jailTime As Byte
        Dim Count    As Byte
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        jailTime = .incomingData.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If
        
        '/carcel nick@motivo@<tiempo>
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If EsGM(tUser) Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")

                        End If

                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")

                        End If
                        
                        If PersonajeExiste(UserName) Then
                            If Database_Enabled Then
                                Call SavePenaDatabase(UserName, .name & ": CARCEL " & jailTime & "m, MOTIVO: " & Reason & " " & Date & " " & Time)
                            Else
                                Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)

                            End If

                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name)
                        Call LogGM(.name, " encarceló a " & UserName)

                    End If

                End If

            End If

        End If

    End With
        
    Exit Sub
        
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "KillNPC" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
        
    On Error GoTo HandleKillNPC_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
    
        
        
        
        If .flags.Privilegios And PlayerType.user Then Exit Sub

        'Si estamos en el mapa pretoriano...
        If .Pos.Map = MAPA_PRETORIANO Then

            '... solo los Dioses y Administradores pueden usar este comando en el mapa pretoriano.
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then
                
                Call WriteConsoleMsg(UserIndex, "Solo los Administradores y Dioses pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        Dim tNPC As Integer
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then

            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & NpcList(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
            Dim auxNPC As npc
            auxNPC = NpcList(tNPC)
            
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
        Else
            Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleKillNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleKillNPC", Erl)

    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "WarnUser" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Reason   As String

        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        
        ' Tenes que ser Admin, Dios o Semi-Dios
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) = 0 Then Exit Sub
        
        ' Me fijo que esten todos los parametros.
        If Len(UserName) = 0 Or Len(Trim$(Reason)) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Formato inválido. /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim tUser As Integer
        tUser = NameIndex(UserName)
        
        ' No advertir a GM's
        If EsGM(tUser) Then
            Call WriteConsoleMsg(UserIndex, "No podes advertir a Game Masters.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If (InStrB(UserName, "\") <> 0) Then
            UserName = Replace(UserName, "\", "")

        End If

        If (InStrB(UserName, "/") <> 0) Then
            UserName = Replace(UserName, "/", "")

        End If
                    
        If PersonajeExiste(UserName) Then

            If Database_Enabled Then
                Call SaveWarnDatabase(UserName, "ADVERTENCIA: " & Reason & " " & Date & " " & Time, .name)
 
            Else
                
                Dim Count As Integer
                Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)

            End If
            
            ' Para el GM
            Call WriteConsoleMsg(UserIndex, "Has advertido a " & UserName, FontTypeNames.FONTTYPE_CENTINELA)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " ha advertido a " & UserName & " por " & Reason, FontTypeNames.FONTTYPE_GM))
            Call LogGM(.name, " advirtio a " & UserName & " por " & Reason)

            ' Si esta online...
            If tUser >= 0 Then
                ' Actualizo el valor en la memoria.
                UserList(tUser).Stats.Advertencias = UserList(tUser).Stats.Advertencias + 1
                
                ' Para el usuario advertido
                Call WriteConsoleMsg(tUser, "Has sido advertido por " & Reason, FontTypeNames.FONTTYPE_CENTINELA)
                Call WriteConsoleMsg(tUser, "Tenés " & UserList(tUser).Stats.Advertencias & " advertencias actualmente.", FontTypeNames.FONTTYPE_CENTINELA)
                
                ' Cuando acumulas cierta cantidad de advertencias...
                Select Case UserList(tUser).Stats.Advertencias
                
                    Case 3
                        Call Encarcelar(tUser, 30, "Servidor")
                    
                    Case 5
                        ' TODO: Banear PJ alv.
                    
                End Select
                
            End If

        End If
        
    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWarnUser", Erl)

    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleMensajeUser(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Mensaje  As String
        Dim privs    As PlayerType
        Dim Count    As Byte
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        Mensaje = .incomingData.ReadASCIIString()
        
        tUser = NameIndex(UserName)
        
        If EsGM(UserIndex) Then
        
            If LenB(UserName) = 0 Or LenB(Mensaje) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /MENSAJEINFORMACION nick@mensaje", FontTypeNames.FONTTYPE_INFO)
                
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")

                End If
                    
                AddCorreo UserIndex, UserName, LCase$(Mensaje), 0, 0

            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMensajeUser", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleTraerBoveda(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jul/2014
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Call UpdateUserHechizos(True, UserIndex, 0)
       
        Call UpdateUserInv(True, UserIndex, 0)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTraerBoveda", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/28/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName      As String
        Dim tUser         As Integer

        Dim opcion        As Byte
        Dim Arg1          As String
        Dim Arg2          As String

        Dim valido        As Boolean

        Dim LoopC         As Byte
        Dim commandString As String
        Dim n             As Byte
        Dim tmpLong       As Long
        
        UserName = Replace(.incomingData.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
            
        Else
            tUser = NameIndex(UserName)

        End If
        
        opcion = .incomingData.ReadByte()
        Arg1 = .incomingData.ReadASCIIString()
        Arg2 = .incomingData.ReadASCIIString()

        ' Si no es GM, no hacemos nada.
        If Not EsGM(UserIndex) Then Exit Sub
        
        ' Si NO sos Dios o Admin,
        If (.flags.Privilegios And PlayerType.Admin) = 0 Then

            ' Si te editas a vos mismo esta bien ;)
            If UserIndex <> tUser Then Exit Sub
            
        End If
        
        Select Case opcion

            Case eEditOptions.eo_Gold

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    UserList(tUser).Stats.GLD = val(Arg1)
                    Call WriteUpdateGold(tUser)

                End If
                
            Case eEditOptions.eo_Experience

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If UserList(tUser).Stats.ELV < STAT_MAXELV Then
                        UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                        Call CheckUserLevel(tUser)
                        Call WriteUpdateExp(tUser)
                            
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario es nivel máximo.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
            Case eEditOptions.eo_Body

                If tUser <= 0 Then
                    
                    If Database_Enabled Then
                        Call SaveUserBodyDatabase(UserName, val(Arg1))
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Body", Arg1)

                    End If

                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                End If
                   
            Case eEditOptions.eo_Arma
                
                If tUser <= 0 Then
                       
                    If Database_Enabled Then
                        'Call SaveUserBodyDatabase(UserName, val(Arg1))
                    Else
                        'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, val(Arg1), UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    
                End If
                       
            Case eEditOptions.eo_Escudo
                
                If tUser <= 0 Then
                       
                    If Database_Enabled Then
                        'Call SaveUserBodyDatabase(UserName, val(Arg1))
                    Else
                        'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, val(Arg1), UserList(tUser).Char.CascoAnim)
                    
                End If
                       
            Case eEditOptions.eo_Casco
                
                If tUser <= 0 Then
                       
                    If Database_Enabled Then
                        'Call SaveUserBodyDatabase(UserName, val(Arg1))
                    Else
                        'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, val(Arg1))
                    
                End If
                       
            Case eEditOptions.eo_Particula
                
                If Not .flags.Privilegios = Consejero Then
                    If tUser <= 0 Then

                        If Database_Enabled Then
                            'Call SaveUserBodyDatabase(UserName, val(Arg1))
                        Else
                            'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)

                        End If

                        Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        'Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, val(Arg1))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, val(Arg1), 9999, False))
                        .Char.ParticulaFx = val(Arg1)
                        .Char.loops = 9999

                    End If

                End If
                
            Case eEditOptions.eo_Head

                If tUser <= 0 Then
                    
                    If Database_Enabled Then
                        Call SaveUserHeadDatabase(UserName, val(Arg1))
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Head", Arg1)

                    End If

                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                End If
                
            Case eEditOptions.eo_CriminalsKilled

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > MAXUSERMATADOS Then
                        UserList(tUser).Faccion.CriminalesMatados = MAXUSERMATADOS
                    Else
                        UserList(tUser).Faccion.CriminalesMatados = val(Arg1)

                    End If

                End If
                
            Case eEditOptions.eo_CiticensKilled

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > MAXUSERMATADOS Then
                        UserList(tUser).Faccion.ciudadanosMatados = MAXUSERMATADOS
                    Else
                        UserList(tUser).Faccion.ciudadanosMatados = val(Arg1)

                    End If

                End If
                
            Case eEditOptions.eo_Level

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > STAT_MAXELV Then
                        Arg1 = CStr(STAT_MAXELV)
                        Call WriteConsoleMsg(UserIndex, "No podés tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)

                    End If
                        
                    UserList(tUser).Stats.ELV = val(Arg1)

                End If
                    
                Call WriteUpdateUserStats(UserIndex)
                
            Case eEditOptions.eo_Class

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    For LoopC = 1 To NUMCLASES

                        If Tilde(ListaClases(LoopC)) = Tilde(Arg1) Then Exit For
                    Next LoopC
                        
                    If LoopC > NUMCLASES Then
                        Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(tUser).clase = LoopC

                    End If

                End If
                
            Case eEditOptions.eo_Skills

                For LoopC = 1 To NUMSKILLS

                    If Tilde(Replace$(SkillsNames(LoopC), " ", "+")) = Tilde(Arg1) Then Exit For
                Next LoopC
                    
                If LoopC > NUMSKILLS Then
                    Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                Else

                    If tUser <= 0 Then
                        
                        If Database_Enabled Then
                            Call SaveUserSkillDatabase(UserName, LoopC, val(Arg2))
                        Else
                            Call WriteVar(CharPath & UserName & ".chr", "Skills", "SK" & LoopC, Arg2)

                        End If

                        Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)

                    End If

                End If
                
            Case eEditOptions.eo_SkillPointsLeft

                If tUser <= 0 Then
                    
                    If Database_Enabled Then
                        Call SaveUserSkillsLibres(UserName, val(Arg1))
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "STATS", "SkillPtsLibres", Arg1)

                    End If
                        
                    Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    UserList(tUser).Stats.SkillPts = val(Arg1)

                End If
                
            Case eEditOptions.eo_Sex

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Arg1 = UCase$(Arg1)

                    If (Arg1 = "MUJER") Then
                        UserList(tUser).genero = eGenero.Mujer
                        
                    ElseIf (Arg1 = "HOMBRE") Then
                        UserList(tUser).genero = eGenero.Hombre

                    End If

                End If
                
            Case eEditOptions.eo_Raza

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        
                Else
                    
                    Arg1 = UCase$(Arg1)

                    If (Arg1 = "HUMANO") Then
                        UserList(tUser).raza = eRaza.Humano
                            
                    ElseIf (Arg1 = "ELFO") Then
                        UserList(tUser).raza = eRaza.Elfo
                            
                    ElseIf (Arg1 = "DROW") Then
                        UserList(tUser).raza = eRaza.Drow
                            
                    ElseIf (Arg1 = "ENANO") Then
                        UserList(tUser).raza = eRaza.Enano
                            
                    ElseIf (Arg1 = "GNOMO") Then
                        UserList(tUser).raza = eRaza.Gnomo
                            
                    ElseIf (Arg1 = "ORCO") Then
                        UserList(tUser).raza = eRaza.Orco

                    End If

                End If
                    
            Case eEditOptions.eo_Vida

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong > 0 Then
                        UserList(tUser).Stats.MaxHp = Min(tmpLong, STAT_MAXHP)
                        UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MaxHp
                            
                        Call WriteUpdateUserStats(tUser)

                    End If

                End If
                    
            Case eEditOptions.eo_Mana

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong > 0 Then
                        UserList(tUser).Stats.MaxMAN = Min(tmpLong, STAT_MAXMP)
                        UserList(tUser).Stats.MinMAN = UserList(tUser).Stats.MaxMAN
                            
                        Call WriteUpdateUserStats(tUser)

                    End If

                End If
                    
            Case eEditOptions.eo_Energia

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong > 0 Then
                        UserList(tUser).Stats.MaxSta = Min(tmpLong, STAT_MAXSTA)
                        UserList(tUser).Stats.MinSta = UserList(tUser).Stats.MaxSta
                            
                        Call WriteUpdateUserStats(tUser)

                    End If

                End If
                        
            Case eEditOptions.eo_MinHP

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong >= 0 Then
                        UserList(tUser).Stats.MinHp = Min(tmpLong, STAT_MAXHP)
                            
                        Call WriteUpdateHP(tUser)

                    End If

                End If
                    
            Case eEditOptions.eo_MinMP

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong >= 0 Then
                        UserList(tUser).Stats.MinMAN = Min(tmpLong, STAT_MAXMP)
                            
                        Call WriteUpdateMana(tUser)

                    End If

                End If
                    
            Case eEditOptions.eo_Hit

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong >= 0 Then
                        UserList(tUser).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
                        UserList(tUser).Stats.MinHIT = UserList(tUser).Stats.MaxHit

                    End If

                End If
                    
            Case eEditOptions.eo_MinHit

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong >= 0 Then
                        UserList(tUser).Stats.MinHIT = Min(tmpLong, STAT_MAXHIT)

                    End If

                End If
                    
            Case eEditOptions.eo_MaxHit

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    tmpLong = val(Arg1)
                        
                    If tmpLong >= 0 Then
                        UserList(tUser).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)

                    End If

                End If
                    
            Case eEditOptions.eo_Desc

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                ElseIf DescripcionValida(Arg1) Then
                    UserList(tUser).Desc = Arg1
                        
                Else
                    Call WriteConsoleMsg(UserIndex, "Caracteres inválidos en la descripción.", FontTypeNames.FONTTYPE_INFO)

                End If
                    
            Case eEditOptions.eo_Intervalo

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Arg1 = UCase$(Arg1)
                        
                    tmpLong = val(Arg2)
                        
                    If tmpLong >= 0 Then
                    
                        Select Case Arg1

                            Case "USAR"
                                UserList(tUser).Intervals.UsarClic = tmpLong
                                UserList(tUser).Intervals.UsarU = tmpLong
                                    
                            Case "USAR_U", "USAR+U", "USAR-U"
                                UserList(tUser).Intervals.UsarU = tmpLong
                                    
                            Case "USAR_CLIC", "USAR+CLIC", "USAR-CLIC", "USAR_CLICK", "USAR+CLICK", "USAR-CLICK"
                                UserList(tUser).Intervals.UsarClic = tmpLong
                                    
                            Case "ARCO", "PROYECTILES"
                                UserList(tUser).Intervals.Arco = tmpLong
                                    
                            Case "GOLPE", "GOLPES", "GOLPEAR"
                                UserList(tUser).Intervals.Golpe = tmpLong
                                    
                            Case "MAGIA", "HECHIZO", "HECHIZOS", "LANZAR"
                                UserList(tUser).Intervals.Magia = tmpLong

                            Case "COMBO"
                                UserList(tUser).Intervals.GolpeMagia = tmpLong
                                UserList(tUser).Intervals.MagiaGolpe = tmpLong

                            Case "GOLPE-MAGIA", "GOLPE-HECHIZO"
                                UserList(tUser).Intervals.GolpeMagia = tmpLong

                            Case "MAGIA-GOLPE", "HECHIZO-GOLPE"
                                UserList(tUser).Intervals.MagiaGolpe = tmpLong
                                    
                            Case "GOLPE-USAR"
                                UserList(tUser).Intervals.GolpeUsar = tmpLong
                                    
                            Case "TRABAJAR", "WORK", "TRABAJO"
                                UserList(tUser).Intervals.TrabajarConstruir = tmpLong
                                UserList(tUser).Intervals.TrabajarExtraer = tmpLong
                                    
                            Case "TRABAJAR_EXTRAER", "EXTRAER", "TRABAJO_EXTRAER"
                                UserList(tUser).Intervals.TrabajarExtraer = tmpLong
                                    
                            Case "TRABAJAR_CONSTRUIR", "CONSTRUIR", "TRABAJO_CONSTRUIR"
                                UserList(tUser).Intervals.TrabajarConstruir = tmpLong
                                    
                            Case Else
                                Exit Sub

                        End Select
                            
                        Call WriteIntervals(tUser)
                            
                    End If

                End If
                    
            Case eEditOptions.eo_Hogar

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Arg1 = UCase$(Arg1)
                    
                    Select Case Arg1

                        Case "NIX"
                            UserList(tUser).Hogar = eCiudad.cNix

                        Case "ULLA", "ULLATHORPE"
                            UserList(tUser).Hogar = eCiudad.cUllathorpe

                        Case "BANDER", "BANDERBILL"
                            UserList(tUser).Hogar = eCiudad.cBanderbill

                        Case "LINDOS"
                            UserList(tUser).Hogar = eCiudad.cLindos

                        Case "ARGHAL"
                            UserList(tUser).Hogar = eCiudad.cArghal

                        Case "ARKHEIN"
                            UserList(tUser).Hogar = eCiudad.cArkhein

                    End Select

                End If
                
            Case Else
                
                Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)

        End Select

        'Log it!
        commandString = "/MOD "
        
        Select Case opcion

            Case eEditOptions.eo_Gold
                commandString = commandString & "ORO "
            
            Case eEditOptions.eo_Experience
                commandString = commandString & "EXP "
            
            Case eEditOptions.eo_Body
                commandString = commandString & "BODY "
            
            Case eEditOptions.eo_Head
                commandString = commandString & "HEAD "
            
            Case eEditOptions.eo_CriminalsKilled
                commandString = commandString & "CRI "
            
            Case eEditOptions.eo_CiticensKilled
                commandString = commandString & "CIU "
            
            Case eEditOptions.eo_Level
                commandString = commandString & "LEVEL "
            
            Case eEditOptions.eo_Class
                commandString = commandString & "CLASE "
            
            Case eEditOptions.eo_Skills
                commandString = commandString & "SKILLS "
            
            Case eEditOptions.eo_SkillPointsLeft
                commandString = commandString & "SKILLSLIBRES "
                
            Case eEditOptions.eo_Sex
                commandString = commandString & "SEX "
                
            Case eEditOptions.eo_Raza
                commandString = commandString & "RAZA "

            Case eEditOptions.eo_Vida
                commandString = commandString & "VIDA "
                    
            Case eEditOptions.eo_Mana
                commandString = commandString & "MANA "
                    
            Case eEditOptions.eo_Energia
                commandString = commandString & "ENERGIA "
                    
            Case eEditOptions.eo_MinHP
                commandString = commandString & "MINHP "
                    
            Case eEditOptions.eo_MinMP
                commandString = commandString & "MINMP "
                    
            Case eEditOptions.eo_Hit
                commandString = commandString & "HIT "
                    
            Case eEditOptions.eo_MinHit
                commandString = commandString & "MINHIT "
                    
            Case eEditOptions.eo_MaxHit
                commandString = commandString & "MAXHIT "
                    
            Case eEditOptions.eo_Desc
                commandString = commandString & "DESC "
                    
            Case eEditOptions.eo_Intervalo
                commandString = commandString & "INTERVALO "
                    
            Case eEditOptions.eo_Hogar
                commandString = commandString & "HOGAR "
                
            Case Else
                commandString = commandString & "UNKOWN "

        End Select
        
        commandString = commandString & Arg1 & " " & Arg2
        
        Call LogGM(.name, commandString & " " & UserName)

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        Dim targetName  As String
        Dim targetIndex As Integer
        
        targetName = Replace$(.incomingData.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
            If targetIndex <= 0 Then

                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)

                End If

            Else

                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)

                End If

            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer

        UserName = .incomingData.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
                
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                Call SendUserInvTxtFromChar(UserIndex, UserName)
                
            Else
                Call SendUserInvTxt(UserIndex, tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Long
        Dim message  As String
        
        UserName = .incomingData.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & UserName)
            
            If tUser <= 0 Then
            
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")

                End If
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ReviveChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                With UserList(tUser)

                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        'Call DarCuerpoDesnudo(tUser)
                        
                        'Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Call RevivirUsuario(tUser)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp

                End With
                
                ' Call WriteHora(tUser)
                Call WriteUpdateHP(tUser)
                Call ActualizarVelocidadDeUsuario(tUser)
                Call LogGM(.name, "Resucito a " & UserName)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "OnlineGM" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOnlineGM_Err

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 12/28/06
    '
    '***************************************************
    Dim i    As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
         
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).name & ", "

            End If

        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleOnlineGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOnlineGM", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOnlineMap_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
    
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        Dim LoopC As Long
        Dim list  As String
        Dim priv  As PlayerType
        
        priv = PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser

            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.Map = .Pos.Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).name & ", "

            End If

        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleOnlineMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOnlineMap", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Forgive" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)
        
    On Error GoTo HandleForgive_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
  
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate NPC and make sure player is not dead
        If (NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (NpcList(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        
        Dim priest As npc
        priest = NpcList(.flags.TargetNPC)

        'Make sure it's close enough
        If Distancia(.Pos, priest.Pos) > 3 Then
            'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .Faccion.Status = 1 Or .Faccion.ArmadaReal = 1 Then
            'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            Call WriteChatOverHead(UserIndex, "Tu alma ya esta libre de pecados hijo mio.", priest.Char.CharIndex, vbWhite)
            Exit Sub

        End If
        
        If .Faccion.FuerzasCaos > 0 Then
            Call WriteChatOverHead(UserIndex, "¡¡Dios no te perdonará mientras seas fiel al Demonio!!", priest.Char.CharIndex, vbWhite)
            Exit Sub

        End If

        If .GuildIndex <> 0 Then
            If modGuilds.Alineacion(.GuildIndex) = 1 Then
                Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.", priest.Char.CharIndex, vbWhite)
                Exit Sub

            End If

        End If

        If .Faccion.ciudadanosMatados > 0 Then
            Dim Donacion As Long
            Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano

            Call WriteChatOverHead(UserIndex, "Has matado a ciudadanos inocentes, Dios no puede perdonarte lo que has hecho. " & "Pero si haces una generosa donación de, digamos, " & PonerPuntos(Donacion) & " monedas de oro, tal vez cambie de opinión...", priest.Char.CharIndex, vbWhite)
            Exit Sub

        End If

        Call WriteChatOverHead(UserIndex, "Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", priest.Char.CharIndex, vbYellow)

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, "80", 100, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call VolverCiudadano(UserIndex)

    End With
        
    Exit Sub

HandleForgive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForgive", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Kick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        Dim rank     As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                
            Else

                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.name, "Echo a " & UserName)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Execute" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
 
                Call UserDie(tUser)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserList(tUser).name, FontTypeNames.FONTTYPE_EJECUCION))
                Call LogGM(.name, " ejecuto a " & UserName)

            Else
            
                Call WriteConsoleMsg(UserIndex, "No está online", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "BanChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
  
        Dim UserName As String
        Dim Reason   As String
        
        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call BanCharacter(UserIndex, UserName, Reason)
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "UnbanChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
            UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else

                If ObtenerBaneo(UserName) Then
                    Call UnBan(UserName)
                    
                    If Database_Enabled Then
                        Call SavePenaDatabase(UserName, .name & ": UNBAN. " & Date & " " & Time)
                    Else

                        'penas
                        Dim cantPenas As Byte

                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, .name & ": UNBAN. " & Date & " " & Time)

                    End If

                    Call LogGM(.name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " desbaneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "NPCFollow" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
        
    On Error GoTo HandleNPCFollow_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
        
            Call DoFollow(.flags.TargetNPC, .name)
            
            NpcList(.flags.TargetNPC).flags.Inmovilizado = 0
            NpcList(.flags.TargetNPC).flags.Paralizado = 0
            NpcList(.flags.TargetNPC).Contadores.Paralisis = 0

        End If

    End With
        
    Exit Sub

HandleNPCFollow_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCFollow", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "SummonChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
            
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If LenB(UserName) <> 0 Then
                tUser = NameIndex(UserName)

                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
            ElseIf .flags.TargetUser > 0 Then
                tUser = .flags.TargetUser

                ' Mover NPCs
            ElseIf .flags.TargetNPC > 0 Then

                If NpcList(.flags.TargetNPC).Pos.Map = .Pos.Map Then
                    Call WarpNpcChar(.flags.TargetNPC, .Pos.Map, .Pos.X, .Pos.Y + 1, True)
                    Call WriteConsoleMsg(UserIndex, "Has desplazado a la criatura.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Sólo puedes mover NPCs dentro del mismo mapa.", FontTypeNames.FONTTYPE_INFO)

                End If

                Exit Sub

            Else
                Exit Sub

            End If

            If CompararPrivilegiosUser(tUser, UserIndex) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser).name & " que quieres traerlo a tu posición.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " quiere transportarte a su ubicación. Escribe /ira " & .name & " para ir.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
            Dim NotConsejero As Boolean
            NotConsejero = (.flags.Privilegios And PlayerType.Consejero) = 0
                
            ' Consejeros sólo pueden traer en el mismo mapa
            If NotConsejero Or .Pos.Map = UserList(tUser).Pos.Map Then
                    
                    ' Si el admin está invisible no mostramos el nombre
148                 If NotConsejero And .flags.AdminInvisible = 1 Then
150                     Call WriteConsoleMsg(tUser, "Te han trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Else
152                     Call WriteConsoleMsg(tUser, .name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    'HarThaoS: Si lo sumonean a un mapa interdimensional desde uno no interdimensional me guardo la posición de donde viene.
                    If EsMapaInterdimensional(.Pos.Map) And Not EsMapaInterdimensional(UserList(tUser).Pos.Map) Then
                        UserList(tUser).flags.ReturnPos = UserList(tUser).Pos
                    End If
                    
                    

                Call WarpToLegalPos(tUser, .Pos.Map, .Pos.X, .Pos.Y + 1, True, True)

                Call WriteConsoleMsg(UserIndex, "Has traído a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                    
                Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                
            End If
            
        End If

    End With

    Exit Sub
        
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSummonChar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSpawnListRequest_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)

    End With
        
    Exit Sub

HandleSpawnListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpawnListRequest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSpawnCreature_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim npc As Integer
            npc = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then
                Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            End If
            
            Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)

        End If

    End With
        
    Exit Sub

HandleSpawnCreature_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpawnCreature", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
        
    On Error GoTo HandleResetNPCInventory_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.name, "/RESETINV " & NpcList(.flags.TargetNPC).name)

    End With
        
    Exit Sub

HandleResetNPCInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleResetNPCInventory", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCleanWorld_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then Exit Sub

        Call LimpiezaForzada
            
        Call WriteConsoleMsg(UserIndex, "Se han limpiado los items del suelo.", FontTypeNames.FONTTYPE_INFO)
            
    End With

    Exit Sub

HandleCleanWorld_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCleanWorld", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            If LenB(message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & message, FontTypeNames.FONTTYPE_SERVER))

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleServerMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "NickToIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 24/07/07
    'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        Dim priv     As PlayerType
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        
            tUser = NameIndex(UserName)
            Call LogGM(.name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
                
            Else
                priv = PlayerType.user

            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)

                    Dim ip    As String
                    Dim lista As String
                    Dim LoopC As Long

                    ip = UserList(tUser).ip

                    For LoopC = 1 To LastUser

                        If UserList(LoopC).ip = ip Then
                        
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                            
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).name & ", "
                                End If

                            End If

                        End If

                    Next LoopC

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else
            
                Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "IPToNick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
        
    On Error GoTo HandleIPToNick_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim ip    As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv  As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            
        Else
            priv = PlayerType.user

        End If

        For LoopC = 1 To LastUser

            If UserList(LoopC).ip = ip Then
            
                If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).name & ", "
                    End If

                End If

            End If

        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleIPToNick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleIPToNick", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim GuildName As String
        Dim tGuild    As Integer
        
        GuildName = .incomingData.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
        
    On Error GoTo HandleTeleportCreate_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte
        
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Call LogGM(.name, "/CT " & Mapa & "," & X & "," & Y)
        
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No podés crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim Objeto As obj
        
        Objeto.amount = 1
        Objeto.ObjIndex = 378
        
        Call MakeObj(Objeto, .Pos.Map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With

    End With
        
    Exit Sub

HandleTeleportCreate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTeleportCreate", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
        
    On Error GoTo HandleTeleportDestroy_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte

        '/dt
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            'Si no tengo objeto y no tengo traslado
            If .ObjInfo.ObjIndex = 0 And .TileExit.Map = 0 Then Exit Sub
                
            'Si no tengo objeto pero tengo traslado
            If .ObjInfo.ObjIndex = 0 And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & X & "," & Y)
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
                
                'si tengo objeto y traslado
            ElseIf .ObjInfo.ObjIndex > 0 And ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With

    End With
        
    Exit Sub

HandleTeleportDestroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTeleportDestroy", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RainToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRainToggle_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        Call LogGM(.name, "/LLUVIA")
        
        Lloviendo = Not Lloviendo
        Nebando = Not Nebando
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

        If Lloviendo Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
            Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HF5D3F3, 250)) 'Rayo
            Call ApagarFogatas

        End If

    End With
        
    Exit Sub

HandleRainToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim tUser As Integer
        Dim Desc  As String
        
        Desc = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
        
            tUser = .flags.TargetUser

            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
                
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
        
    On Error GoTo HanldeForceMIDIToMap_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)
    
        Dim midiID As Byte
        Dim Mapa   As Integer
        
        midiID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).music_numberLow))
                
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))

            End If

        End If

    End With
        
    Exit Sub

HanldeForceMIDIToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HanldeForceMIDIToMap", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
        
    On Error GoTo HandleForceWAVEToMap_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim waveID As Byte
        Dim Mapa   As Integer
        Dim X      As Byte
        Dim Y      As Byte
        
        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
            
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y

            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))

        End If

    End With
        
    Exit Sub

HandleForceWAVEToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceWAVEToMap", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster)) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
                
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDestroyAllItemsInArea_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
  
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                    
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If

                    End If

                End If

            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).name, "/MASSDEST")

    End With
        
    Exit Sub

HandleDestroyAllItemsInArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDestroyAllItemsInArea", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legión Oscura.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
        
    On Error GoTo HandleItemsInTheFloor_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Dim tObj  As Integer
        Dim lista As String
        Dim X     As Long
        Dim Y     As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex

                If tObj > 0 Then
                
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)
                    End If

                End If

            Next Y
        Next X

    End With
        
    Exit Sub

HandleItemsInTheFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleItemsInTheFloor", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
                
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call WriteDumb(tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMakeDumb", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call WriteDumbNoMore(tUser)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMakeDumbNoMore", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDumpIPTables_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Call SecurityIp.DumpTables

    End With
        
    Exit Sub

HandleDumpIPTables_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumpIPTables", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
  
    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                If PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos", FontTypeNames.FONTTYPE_INFO)
                    
                    If Database_Enabled Then
                        Call EcharConsejoDatabase(UserName)
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                        Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No existe el personaje.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))

                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de la Legiín Oscura", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legión Oscura", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCouncilKick", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "SetTrigger" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSetTrigger_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
 
        Dim tTrigger As Byte
        Dim tLog     As String
        
        tTrigger = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        If tTrigger >= 0 Then
        
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.name, tLog)
            
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleSetTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetTrigger", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
        
    On Error GoTo HandleAskTrigger_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    '
    '***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleAskTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAskTrigger", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBannedIPList_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Dim lista As String

        Dim LoopC As Long
        
        Call LogGM(.name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleBannedIPList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBannedIPList", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBannedIPReload_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar

    End With
        
    Exit Sub

HandleBannedIPReload_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBannedIPReload", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "GuildBan" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim GuildName   As String
        Dim cantMembers As Integer
        Dim LoopC       As Long
        Dim member      As String
        Dim Count       As Byte
        Dim tIndex      As Integer
        Dim tFile       As String
        
        GuildName = .incomingData.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    'member es la victima
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)

                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)

                    End If
                    
                    If Database_Enabled Then
                        Call SaveBanDatabase(member, .name & " - BAN AL CLAN: " & GuildName & ". " & Date & " " & Time, .name)
                        
                    Else
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & member & ".chr", "BAN", "BANEADO", "1")
                        Call WriteVar(CharPath & member & ".chr", "BAN", "BannedBy", .name)
                        Call WriteVar(CharPath & member & ".chr", "BAN", "BanMotivo", "clan baneado")
                        'ponemos la pena
                        Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, .name & " - BAN AL CLAN: " & GuildName & ". " & Date & " " & Time)

                    End If

                Next LoopC

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "BanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim bannedIP As String
        Dim tUser    As Integer
        Dim Reason   As String
        Dim i        As Long
        
        ' Is it by ip??
        If .incomingData.ReadBoolean() Then
            bannedIP = .incomingData.ReadByte() & "."
            bannedIP = bannedIP & .incomingData.ReadByte() & "."
            bannedIP = bannedIP & .incomingData.ReadByte() & "."
            bannedIP = bannedIP & .incomingData.ReadByte()
            
        Else
        
            tUser = NameIndex(.incomingData.ReadASCIIString())
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
                
            Else
                bannedIP = UserList(tUser).ip

            End If

        End If
        
        Reason = .incomingData.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
        
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser

                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call BanCharacter(UserIndex, UserList(i).name, "IP POR " & Reason)

                        End If

                    End If

                Next i

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
        
    On Error GoTo HandleUnbanIP_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
        
    Exit Sub

HandleUnbanIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUnbanIP", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CreateItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCreateItem_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim tObj    As Integer
        Dim Cuantos As Integer
        
        tObj = .incomingData.ReadInteger()
        Cuantos = .incomingData.ReadInteger()
    
        ' Si es usuario, lo sacamos cagando.
        If Not EsGM(UserIndex) Then Exit Sub
        
        ' Si es Semi-Dios, dejamos crear un item siempre y cuando pueda estar en el inventario.
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 And ObjData(tObj).Agarrable = 1 Then Exit Sub

        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > MAX_INVENTORY_OBJS Then
            Call WriteConsoleMsg(UserIndex, "Solo podés crear hasta " & CStr(MAX_INVENTORY_OBJS) & " unidades", FontTypeNames.FONTTYPE_TALK)
            Exit Sub

        End If
        
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
        ' El nombre del objeto es nulo?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Objeto As obj
        Objeto.amount = Cuantos
        Objeto.ObjIndex = tObj

        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demas objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(UserIndex, Objeto) Then
                Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.amount & " unidades de " & ObjData(tObj).name & ".", FontTypeNames.FONTTYPE_INFO)
            
            Else

                Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si no hay espacio y es Dios o Admin, lo tiro al piso.
                If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
                    Call TirarItemAlPiso(.Pos, Objeto)
                    Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)

                End If
                
            End If
        
        Else
        
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            ' Si no hay espacio y es Dios o Admin, lo tiro al piso.
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)

            End If

        End If
        
        ' Lo registro en los logs.
        Call LogGM(.name, "/CI: " & tObj & " Cantidad : " & Cuantos)

    End With
        
    Exit Sub

HandleCreateItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDestroyItems_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.name, "/DEST")

        Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, .Pos.X, .Pos.Y)

    End With
        
    Exit Sub

HandleDestroyItems_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDestroyItems", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If

            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                UserList(tUser).Faccion.FuerzasCaos = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                
            Else

                If PersonajeExiste(UserName) Then
                    If Database_Enabled Then
                        Call EcharLegionDatabase(UserName)
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .name)

                    End If
                    
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If

            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
            
                UserList(tUser).Faccion.ArmadaReal = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                
            Else

                If PersonajeExiste(UserName) Then
                
                    If Database_Enabled Then
                        Call EcharArmadaDatabase(UserName)
                        
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                        Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .name)

                    End If

                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
        
    On Error GoTo HandleForceMIDIAll_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim midiID As Byte
            midiID = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))

    End With
        
    Exit Sub

HandleForceMIDIAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceMIDIAll", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
        
    On Error GoTo HandleForceWAVEAll_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim waveID As Byte
            waveID = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))

    End With
        
    Exit Sub

HandleForceWAVEAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceWAVEAll", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 1/05/07
    'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName   As String
        Dim punishment As Byte
        Dim NewText    As String
        
        UserName = .incomingData.ReadASCIIString()
        punishment = .incomingData.ReadByte
        NewText = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
                
            Else

                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")

                End If
                
                If PersonajeExiste(UserName) Then
                    Call LogGM(.name, "Borro la pena " & punishment & " de " & UserName & " y la cambií por: " & NewText)
                    
                    If Database_Enabled Then
                        Call CambiarPenaDatabase(UserName, punishment, .name & ": <" & NewText & "> " & Date & " " & Time)
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, .name & ": <" & NewText & "> " & Date & " " & Time)

                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleTileBlockedToggle_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub

        Call LogGM(.name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = eBlock.ALL_SIDES Or eBlock.GM
            
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, IIf(MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked > 0, eBlock.ALL_SIDES, 0))

    End With
        
    Exit Sub

HandleTileBlockedToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTileBlockedToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
        
    On Error GoTo HandleKillNPCNoRespawn_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        
        Call LogGM(.name, "/MATA " & NpcList(.flags.TargetNPC).name)

    End With
        
    Exit Sub

HandleKillNPCNoRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleKillNPCNoRespawn", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
        
    On Error GoTo HandleKillAllNearbyNPCs_Err

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
            
        'Si está en el mapa pretoriano, me aseguro de que los saque correctamente antes que nada.
        If .Pos.Map = MAPA_PRETORIANO Then Call EliminarPretorianos(MAPA_PRETORIANO)

        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then

                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then
                    
                        Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)

                    End If

                End If

            Next X
        Next Y

        Call LogGM(.name, "/MASSKILL")

    End With
        
    Exit Sub

HandleKillAllNearbyNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleKillAllNearbyNPCs", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "LastIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName   As String
        Dim lista      As String
        Dim LoopC      As Byte
        Dim priv       As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then

            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If

            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
                
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If
            
            If validCheck Then
                Call LogGM(.name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectí son:"

                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC

                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLastIP", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ChatColor" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChatColor_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the user`s chat color
    '***************************************************

    With UserList(UserIndex)

        Dim Color As Long
            Color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = Color
        End If

    End With
        
    Exit Sub

HandleChatColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatColor", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Ignored" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
        
    On Error GoTo HandleIgnored_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Ignore the user
    '***************************************************
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If

    End With
        
    Exit Sub

HandleIgnored_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleIgnored", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Check one Users Slot in Particular from Inventory
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim slot     As Byte
        Dim tIndex   As Integer
        
        UserName = .incomingData.ReadASCIIString() 'Que UserName?
        slot = .incomingData.ReadByte() 'Que Slot?
        tIndex = NameIndex(UserName)  'Que user index?

        If Not EsGM(UserIndex) Then Exit Sub
        
        Call LogGM(.name, .name & " Checkeo el slot " & slot & " de " & UserName)
           
        If tIndex > 0 Then
            If slot > 0 And slot <= UserList(UserIndex).CurrentInventorySlots Then
            
                If UserList(tIndex).Invent.Object(slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, " Objeto " & slot & ") " & ObjData(UserList(tIndex).Invent.Object(slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(slot).amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Slot Invílido.", FontTypeNames.FONTTYPE_TALK)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)
        
    On Error GoTo HandleResetAutoUpdate_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reset the AutoUpdate
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleResetAutoUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleResetAutoUpdate", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "Restart" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRestart_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Restart the game
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.name, .name & " reinicio el mundo")
        
        Call ReiniciarServidor(True)

    End With
        
    Exit Sub

HandleRestart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRestart", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
        
    On Error GoTo HandleReloadObjects_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the objects
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado a los objetos.")
        
        Call LoadOBJData
        Call LoadPesca
        Call LoadRecursosEspeciales
        Call WriteConsoleMsg(UserIndex, "Obj.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

    End With
        
    Exit Sub

HandleReloadObjects_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleReloadObjects", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
        
    On Error GoTo HandleReloadSpells_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the spells
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado los hechizos.")
        
        Call CargarHechizos

    End With
        
    Exit Sub

HandleReloadSpells_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleReloadSpells", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
        
    On Error GoTo HandleReloadServerIni_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s INI
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado los INITs.")
        
        Call LoadSini

    End With
        
    Exit Sub

HandleReloadServerIni_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleReloadServerIni", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
        
    On Error GoTo HandleReloadNPCs_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s NPC
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
         
        Call LogGM(.name, .name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

    End With
        
    Exit Sub

HandleReloadNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleReloadNPCs", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "RequestTCPStats" message
' @param UserIndex The index of the user sending the message

Public Sub HandleRequestTCPStats(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRequestTCPStats_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Send the TCP`s stadistics
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
                
        Dim list  As String

        Dim Count As Long

        Dim i     As Long
        
        Call LogGM(.name, .name & " ha pedido las estadisticas del TCP.")
    
        Call WriteConsoleMsg(UserIndex, "Los datos estín en BYTES.", FontTypeNames.FONTTYPE_INFO)
        
        'Send the stats
        With TCPESStats
            Call WriteConsoleMsg(UserIndex, "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando, FontTypeNames.FONTTYPE_INFO)

        End With
        
        'Search for users that are working
        For i = 1 To LastUser

            With UserList(i)

                If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
                    If .outgoingData.Length > 0 Then
                        list = list & .name & " (" & CStr(.outgoingData.Length) & "), "
                        Count = Count + 1

                    End If

                End If

            End With

        Next i
        
        Call WriteConsoleMsg(UserIndex, "Posibles pjs trabados: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, list, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleRequestTCPStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestTCPStats", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "KickAllChars" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
        
    On Error GoTo HandleKickAllChars_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Kick all the chars that are online
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados

    End With
        
    Exit Sub

HandleKickAllChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleKickAllChars", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "Night" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
        
    On Error GoTo HandleNight_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    '
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub

        HoraMundo = GetTickCount()

        Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
    End With
        
    Exit Sub

HandleNight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNight", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "Day" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDay(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDay_Err

    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub

        HoraMundo = GetTickCount() - DuracionDia \ 2

        Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
    End With
        
    Exit Sub

HandleDay_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDay", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "SetTime" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetTime(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSetTime_Err

    With UserList(UserIndex)
        
        

        Dim HoraDia As Long
        HoraDia = .incomingData.ReadLong
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub

        HoraMundo = GetTickCount() - HoraDia
            
        Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
    End With
        
    Exit Sub

HandleSetTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetTime", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandleDonateGold(ByVal UserIndex As Integer)
        
    On Error GoTo handle

    With UserList(UserIndex)
        
        

        Dim Oro As Long
        Oro = .incomingData.ReadLong

        If Oro <= 0 Then Exit Sub

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim priest As npc
        priest = NpcList(.flags.TargetNPC)

        'Validate NPC is an actual priest and the player is not dead
        If (priest.NPCtype <> eNPCType.Revividor And (priest.NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub

        'Make sure it's close enough
        If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 3 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If .Faccion.Status = 1 Or .Faccion.ArmadaReal = 1 Or .Faccion.FuerzasCaos > 0 Or .Faccion.ciudadanosMatados = 0 Then
            Call WriteChatOverHead(UserIndex, "No puedo aceptar tu donación en este momento...", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If

        If .GuildIndex <> 0 Then
            If modGuilds.Alineacion(.GuildIndex) = 1 Then
                Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... no puedo aceptar tu donación.", priest.Char.CharIndex, vbWhite)
                Exit Sub

            End If

        End If

        If .Stats.GLD < Oro Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim Donacion As Long
        Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano

        If Oro < Donacion Then
            Call WriteChatOverHead(UserIndex, "Dios no puede perdonarte si eres una persona avara.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If

        .Stats.GLD = .Stats.GLD - Oro

        Call WriteUpdateGold(UserIndex)

        Call WriteConsoleMsg(UserIndex, "Has donado " & PonerPuntos(Oro) & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)

        Call WriteChatOverHead(UserIndex, "¡Gracias por tu generosa donación! Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbYellow)

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, "80", 100, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call VolverCiudadano(UserIndex)
    
    End With
        
    Exit Sub

handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDonateGold", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandlePromedio(ByVal UserIndex As Integer)
        
    On Error GoTo handle

    With UserList(UserIndex)

        Call WriteConsoleMsg(UserIndex, ListaClases(.clase) & " " & ListaRazas(.raza) & " nivel " & .Stats.ELV & ".", FONTTYPE_INFOBOLD)
            
        Dim Promedio As Double, Vida As Long
        
        Promedio = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
        Vida = 18.5 + ModRaza(.raza).Constitucion / 6 + Promedio * (.Stats.ELV - 1)

        Call WriteConsoleMsg(UserIndex, "Vida esperada: " & Vida & ". Promedio: " & Promedio, FONTTYPE_INFOBOLD)

        Promedio = CalcularPromedioVida(UserIndex)

        Dim Diff As Long, Color As FontTypeNames, Signo As String
            
        Diff = .Stats.MaxHp - Vida
            
        If Diff < 0 Then
            Color = FONTTYPE_PROMEDIO_MENOR
            Signo = "-"

        ElseIf Diff > 0 Then
            Color = FONTTYPE_PROMEDIO_MAYOR
            Signo = "+"

        Else
            Color = FONTTYPE_PROMEDIO_IGUAL
            Signo = "+"
                
        End If

        Call WriteConsoleMsg(UserIndex, "Vida actual: " & .Stats.MaxHp & " (" & Signo & Abs(Diff) & "). Promedio: " & Round(Promedio, 2), Color)

    End With
        
    Exit Sub

handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePromedio", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandleGiveItem(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim ObjIndex As Integer
        Dim Cantidad As Integer
        Dim Motivo   As String
        Dim tIndex   As Integer
        
        UserName = .incomingData.ReadASCIIString()
        ObjIndex = .incomingData.ReadInteger()
        Cantidad = .incomingData.ReadInteger()
        Motivo = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then

            If ObjData(ObjIndex).Agarrable = 1 Then Exit Sub

            If Cantidad > MAX_INVENTORY_OBJS Then Cantidad = MAX_INVENTORY_OBJS

            ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
            If ObjIndex < 1 Or ObjIndex > NumObjDatas Then Exit Sub
            
            ' El nombre del objeto es nulo?
            If LenB(ObjData(ObjIndex).name) = 0 Then Exit Sub

            ' Está online?
            tIndex = NameIndex(UserName)

            If tIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no está conectado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Dim Objeto As obj
            Objeto.amount = Cantidad
            Objeto.ObjIndex = ObjIndex

            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(tIndex, Objeto) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha otorgado a " & UserList(tIndex).name & " " & Cantidad & " " & ObjData(ObjIndex).name & ": " & Motivo, FontTypeNames.FONTTYPE_ROSA))
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario no tiene espacio en el inventario.", FontTypeNames.FONTTYPE_INFO)

            End If

            ' Lo registro en los logs.
            Call LogGM(.name, "/DAR " & UserName & " - Item: " & ObjData(ObjIndex).name & "(" & ObjIndex & ") Cantidad : " & Cantidad)
            Call LogPremios(.name, UserName, ObjIndex, Cantidad, Motivo)
            
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ShowServerForm" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
        
    On Error GoTo HandleShowServerForm_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Show the server form
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click

    End With
        
    Exit Sub

HandleShowServerForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowServerForm", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "CleanSOS" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCleanSOS_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Clean the SOS
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha borrado los SOS")
        
        Call Ayuda.Reset

    End With
        
    Exit Sub

HandleCleanSOS_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCleanSOS", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "SaveChars" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSaveChars_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Save the characters
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado todos los chars")
        
        Call GuardarUsuarios

    End With
        
    Exit Sub

HandleSaveChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSaveChars", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMapInfoBackup_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************

    With UserList(UserIndex)

        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informaciín sobre el BackUp")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).backup_mode = 1
            
        Else
            MapInfo(.Pos.Map).backup_mode = 0

        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).backup_mode)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).backup_mode, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleChangeMapInfoBackup_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoBackup", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMapInfoPK_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    With UserList(UserIndex)

        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informacion sobre si es seguro el mapa.")
        
        MapInfo(.Pos.Map).Seguro = IIf(isMapPk, 1, 0)

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Seguro: " & MapInfo(.Pos.Map).Seguro, FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleChangeMapInfoPK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoPK", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "SINMAGIA", "SININVI", "NOPKS", "NOCIUD".
    '***************************************************
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) <> 0 Then

            Select Case UCase$(tStr)
                
                Case "NEWBIE"
                    MapInfo(.Pos.Map).Newbie = Not MapInfo(.Pos.Map).Newbie
                    Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": Newbie = " & MapInfo(.Pos.Map).Newbie, FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .Pos.Map & ": Newbie = " & MapInfo(.Pos.Map).Newbie)
                        
                Case "SINMAGIA"
                    MapInfo(.Pos.Map).SinMagia = Not MapInfo(.Pos.Map).SinMagia
                    Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": SinMagia = " & MapInfo(.Pos.Map).SinMagia, FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinMagia = " & MapInfo(.Pos.Map).SinMagia)
                        
                Case "NOPKS"
                    MapInfo(.Pos.Map).NoPKs = Not MapInfo(.Pos.Map).NoPKs
                    Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": NoPKs = " & MapInfo(.Pos.Map).NoPKs, FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoPKs = " & MapInfo(.Pos.Map).NoPKs)
                        
                Case "NOCIUD"
                    MapInfo(.Pos.Map).NoCiudadanos = Not MapInfo(.Pos.Map).NoCiudadanos
                    Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos, FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos)
                        
                Case "SININVI"
                    MapInfo(.Pos.Map).SinInviOcul = Not MapInfo(.Pos.Map).SinInviOcul
                    Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": SinInvi = " & MapInfo(.Pos.Map).SinInviOcul, FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinInvi = " & MapInfo(.Pos.Map).SinInviOcul)
                
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'SINMAGIA', 'SININVI', 'NOPKS', 'NOCIUD'", FontTypeNames.FONTTYPE_INFO)

            End Select

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMapInfoNoMagic_Err

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'MagiaSinEfecto -> Options: "1" , "0".
    '***************************************************

    Dim nomagic As Boolean
    
    With UserList(UserIndex)
  
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
        End If

    End With
        
    Exit Sub

HandleChangeMapInfoNoMagic_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoMagic", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMapInfoNoInvi_Err

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'InviSinEfecto -> Options: "1", "0"
    '***************************************************
    Dim noinvi As Boolean
    
    With UserList(UserIndex)

        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
        End If

    End With
        
    Exit Sub

HandleChangeMapInfoNoInvi_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoInvi", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMapInfoNoResu_Err

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'ResuSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)

        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
        End If

    End With
        
    Exit Sub

HandleChangeMapInfoNoResu_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoResu", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************

    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
            
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).terrain = tStr
                
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).terrain, FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
            
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).zone = tStr
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).zone, FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'DUNGEON' ya que al ingresarlo, NO se sentirí el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "SaveMap" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
        
    On Error GoTo HandleSaveMap_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Saves the map
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleSaveMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSaveMap", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim guild As String
            guild = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "DoBackUp" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
        
    On Error GoTo HandleDoBackUp_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub

        Call LogGM(.name, .name & " ha hecho un backup")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete

    End With
        
    Exit Sub

HandleDoBackUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "AlterName" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user name
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        'Reads the userName and newUser Packets
        Dim UserName     As String
        Dim newName      As String
        Dim changeNameUI As Integer
        Dim GuildIndex   As Integer
        
        UserName = .incomingData.ReadASCIIString()
        newName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
                
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "Baneado", "1")
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", "BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & Time)
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", .name)

                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & Time)
                                
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlterName", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "AlterName" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim newMail  As String
        
        UserName = .incomingData.ReadASCIIString()
        newMail = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
                
            Else

                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call LogGM(.name, "Le ha cambiado el mail a " & UserName)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlterMail", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "AlterPassword" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(.incomingData.ReadASCIIString(), "+", " ")
        copyFrom = Replace(.incomingData.ReadASCIIString(), "+", " ")
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
        
            Call LogGM(.name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " cambiado a: " & Password, FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlterPassword", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCreateNPC_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim NpcIndex As Integer
        NpcIndex = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        'Nos fijamos si es pretoriano.
        If NpcList(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearPretoianos MAPA X Y.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo a " & NpcList(NpcIndex).name & " en mapa " & .Pos.Map)

        End If

    End With
        
    Exit Sub

HandleCreateNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateNPC", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "CreateNPCWithRespawn" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCreateNPCWithRespawn_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo con respawn " & NpcList(NpcIndex).name & " en mapa " & .Pos.Map)

        End If

    End With
        
    Exit Sub

HandleCreateNPCWithRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateNPCWithRespawn", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ImperialArmour" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
        
    On Error GoTo HandleImperialArmour_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        
        
        Dim index    As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Select Case index

            Case 1
                ' ArmaduraImperial1 = objindex
            
            Case 2
                ' ArmaduraImperial2 = objindex
            
            Case 3
                ' ArmaduraImperial3 = objindex
            
            Case 4

                ' TunicaMagoImperial = objindex
        End Select

    End With
        
    Exit Sub

HandleImperialArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleImperialArmour", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ChaosArmour" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChaosArmour_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************

    With UserList(UserIndex)

        Dim index    As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        Select Case index

            Case 1
                '   ArmaduraCaos1 = objindex
            
            Case 2
                '   ArmaduraCaos2 = objindex
            
            Case 3
                '   ArmaduraCaos3 = objindex
            
            Case 4

                '  TunicaMagoCaos = objindex
        End Select

    End With
        
    Exit Sub

HandleChaosArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChaosArmour", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "NavigateToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleNavigateToggle_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/12/07
    '
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
            
        Else
            .flags.Navegando = 1

        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)

    End With
        
    Exit Sub

HandleNavigateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleServerOpenToUsersToggle_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        
        
        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster)) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1

        End If

    End With
        
    Exit Sub

HandleServerOpenToUsersToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleServerOpenToUsersToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "TurnOffServer" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleParticipar(ByVal UserIndex As Integer)
        
    On Error GoTo HandleParticipar_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        
        
        
        If Torneo.HayTorneoaActivo = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ningún evento disponible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
                   
        If .flags.EnTorneo Then
            Call WriteConsoleMsg(UserIndex, "Ya estás participando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If .Stats.ELV > Torneo.nivelmaximo Then
            Call WriteConsoleMsg(UserIndex, "El nivel míximo para participar es " & Torneo.nivelmaximo & ".", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If .Stats.ELV < Torneo.NivelMinimo Then
            Call WriteConsoleMsg(UserIndex, "El nivel mínimo para participar es " & Torneo.NivelMinimo & ".", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If .Stats.GLD < Torneo.costo Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro para ingresar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Mage And Torneo.mago = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Cleric And Torneo.clerico = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Warrior And Torneo.guerrero = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Bard And Torneo.bardo = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Assasin And Torneo.asesino = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
   
        If .clase = Druid And Torneo.druido = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Paladin And Torneo.Paladin = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Hunter And Torneo.cazador = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .clase = Trabajador And Torneo.cazador = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
   
        If Torneo.Participantes = Torneo.cupos Then
            Call WriteConsoleMsg(UserIndex, "Los cupos ya estan llenos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
  
        Call ParticiparTorneo(UserIndex)

    End With
        
    Exit Sub

HandleParticipar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticipar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "TurnCriminal" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)

            If tUser > 0 Then Call VolverCriminal(tUser)

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTurnCriminal", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "ResetFactions" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then Call ResetFacciones(tUser)

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleResetFactions", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim UserName   As String
        Dim GuildIndex As Integer
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveCharFromGuild", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "RequestCharMail" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Request user mail
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim mail     As String
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestCharMail", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "SystemMessage" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim message As String
            message = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSystemMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "SetMOTD" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 03/31/07
    'Set the MOTD
    'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
    '   - Fixed a bug that prevented from properly setting the new number of lines.
    '   - Fixed a bug that caused the player to be kicked.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim newMOTD           As String

        Dim auxiliaryString() As String

        Dim LoopC             As Long
        
        newMOTD = .incomingData.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(DatPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(DatPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetMOTD", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handle the "ChangeMOTD" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
        
    On Error GoTo HandleChangeMOTD_Err

    '***************************************************
    'Author: Juan Martín sotuyo Dodero (Maraxus)
    'Last Modification: 12/29/06
    'Change the MOTD
    '***************************************************
    With UserList(UserIndex)
  
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then Exit Sub

        Dim auxiliaryString As String

        Dim LoopC           As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
        
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If

        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)

    End With
        
    Exit Sub

HandleChangeMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMOTD", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "Ping" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
        
    On Error GoTo HandlePing_Err

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)

        Dim Time As Long
        
        Time = .incomingData.ReadLong()
        
        Call WritePong(UserIndex, Time)

    End With
        
    Exit Sub

HandlePing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePing", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
        
    On Error GoTo FlushBuffer_Err

    '***************************************************
    'Sends all data existing in the buffer
    '***************************************************
    
    With UserList(UserIndex).outgoingData

        If .Length = 0 Then Exit Sub
        
        ' Tratamos de enviar los datos.
        Dim ret As Long
            ret = WsApiEnviar(UserIndex, .ReadAll)
    
        ' Si recibimos un error como respuesta de la API, cerramos el socket.
        If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        
            ' Close the socket avoiding any critical error
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)

        End If
        
        Call .Clean
        
    End With
        
    Exit Sub

FlushBuffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.FlushBuffer", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleQuestionGM(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Consulta       As String
        Dim TipoDeConsulta As String

        Consulta = .incomingData.ReadASCIIString()
        TipoDeConsulta = .incomingData.ReadASCIIString()

        If UserList(UserIndex).donador.activo = 1 Then
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta & "-Prioritario")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).name & "(Prioritario).", FontTypeNames.FONTTYPE_SERVER))
            
        Else
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).name & ".", FontTypeNames.FONTTYPE_SERVER))

        End If

        Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)
        
        Call LogConsulta(.name & " (" & TipoDeConsulta & ") " & Consulta)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleOfertaInicial(ByVal UserIndex As Integer)
        
    On Error GoTo HandleOfertaInicial_Err
    
    With UserList(UserIndex)

        Dim Oferta As Long
            Oferta = .incomingData.ReadLong()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                
            Exit Sub

        End If

        If .flags.TargetNPC < 1 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Subastador Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 2 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.Subastando = False Then
            Call WriteChatOverHead(UserIndex, "Oye amigo, tu no podés decirme cual es la oferta inicial.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If
        
        If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.Subastando = True Then
            UserList(UserIndex).Counters.TiempoParaSubastar = 0
            Subasta.OfertaInicial = Oferta
            Subasta.MejorOferta = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " está subastando: " & ObjData(Subasta.ObjSubastado).name & " (Cantidad: " & Subasta.ObjSubastadoCantidad & " ) - con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas. Escribe /OFERTAR (cantidad) para participar.", FontTypeNames.FONTTYPE_SUBASTA))
            .flags.Subastando = False
            Subasta.HaySubastaActiva = True
            Subasta.Subastador = .name
            Subasta.MinutosDeSubasta = 5
            Subasta.TiempoRestanteSubasta = 300
            Call LogearEventoDeSubasta("#################################################################################################################################################################################################")
            Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
            Call LogearEventoDeSubasta(.name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & " y con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas.")
            frmMain.SubastaTimer.Enabled = True
            Call WarpUserChar(UserIndex, 14, 27, 64, True)

            'lalala toda la bola de los timerrr
        End If

    End With
        
    Exit Sub

HandleOfertaInicial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOfertaInicial", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleOfertaDeSubasta(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Oferta   As Long
        Dim ExOferta As Long
        
        Oferta = .incomingData.ReadLong()
        
        If Subasta.HaySubastaActiva = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFOIAO)
            
            Exit Sub

        End If
               
        If Oferta < Subasta.MejorOferta + 100 Then
            Call WriteConsoleMsg(UserIndex, "Debe haber almenos una diferencia de 100 monedas a la ultima oferta!", FontTypeNames.FONTTYPE_INFOIAO)
            
            Exit Sub

        End If
        
        If .name = Subasta.Subastador Then
            Call WriteConsoleMsg(UserIndex, "No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...", FontTypeNames.FONTTYPE_INFOIAO)
            
            Exit Sub

        End If
        
        If .Stats.GLD >= Oferta Then

            'revisar que pasa si el usuario que oferto antes esta offline
            'Devolvemos el oro al usuario que oferto antes...(si es que hubo oferta)
            If Subasta.HuboOferta = True Then
                ExOferta = NameIndex(Subasta.Comprador)
                UserList(ExOferta).Stats.GLD = UserList(ExOferta).Stats.GLD + Subasta.MejorOferta
                Call WriteUpdateGold(ExOferta)

            End If
            
            Subasta.MejorOferta = Oferta
            Subasta.Comprador = .name
            
            .Stats.GLD = .Stats.GLD - Oferta
            Call WriteUpdateGold(UserIndex)
            
            If Subasta.TiempoRestanteSubasta < 60 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .name & " (Ofrece " & PonerPuntos(Oferta) & " monedas de oro) - Tiempo Extendido. Escribe /SUBASTA para mas informaciín.", FontTypeNames.FONTTYPE_SUBASTA))
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta en el ultimo minuto ofreciendo " & PonerPuntos(Oferta) & " monedas.")
                Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta + 30
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .name & " (Ofrece " & PonerPuntos(Oferta) & " monedas de oro). Escribe /SUBASTA para mas informaciín.", FontTypeNames.FONTTYPE_SUBASTA))
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta ofreciendo " & PonerPuntos(Oferta) & " monedas.")
                Subasta.HuboOferta = True
                Subasta.PosibleCancelo = False

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No posees esa cantidad de oro.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleGlobalMessage(ByVal UserIndex As Integer)
        
    Dim TActual     As Long
    Dim ElapsedTime As Long

    TActual = GetTickCount()
    ElapsedTime = TActual - UserList(UserIndex).Counters.MensajeGlobal
                
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim chat As String

        chat = .incomingData.ReadASCIIString()

        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
            'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
        
        ElseIf ElapsedTime < IntervaloMensajeGlobal Then
            Call WriteConsoleMsg(UserIndex, "No puedes escribir mensajes globales tan rápido.", FontTypeNames.FONTTYPE_WARNING)
        
        Else
            UserList(UserIndex).Counters.MensajeGlobal = TActual

            If EstadoGlobal Then
                If LenB(chat) <> 0 Then
                    'Analize chat...
                    Call Statistics.ParseChat(chat)

                    ' WyroX: Foto-denuncias - Push message
                    Dim i As Integer

                    For i = 1 To UBound(.flags.ChatHistory) - 1
                        .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat

                    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[" & .name & "] " & chat, FontTypeNames.FONTTYPE_GLOBAL))

                    'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbBlue & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                End If

            Else
                Call WriteConsoleMsg(UserIndex, "El global se encuentra Desactivado.", FontTypeNames.FONTTYPE_GLOBAL)

            End If

        End If
    
    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Public Sub HandleGlobalOnOff(ByVal UserIndex As Integer)
        
    On Error GoTo HandleGlobalOnOff_Err

    'Author: Pablo Mercavides
    With UserList(UserIndex)

        If Not EsGM(UserIndex) Then Exit Sub
        Call LogGM(.name, "/GLOBAL")
        
        If EstadoGlobal = False Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Chat general habilitado. Escribe" & Chr(34) & "/CONSOLA" & Chr(34) & " o " & Chr(34) & ";" & Chr(34) & " y su mensaje para utilizarlo.", FontTypeNames.FONTTYPE_SERVER))
            EstadoGlobal = True
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Chat General deshabilitado.", FontTypeNames.FONTTYPE_SERVER))
            EstadoGlobal = False

        End If
        
    End With
        
    Exit Sub

HandleGlobalOnOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleIngresarConCuenta(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Dim Version As String
    
    With UserList(UserIndex)

        Dim CuentaEmail    As String
        Dim CuentaPassword As String
        Dim MacAddress     As String
        Dim HDserial       As Long
        Dim MD5            As String
        
        CuentaEmail = .incomingData.ReadASCIIString()
        CuentaPassword = .incomingData.ReadASCIIString()
        Version = CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte())
        MacAddress = .incomingData.ReadASCIIString()
        HDserial = .incomingData.ReadLong()
        MD5 = .incomingData.ReadASCIIString()
            
        #If DEBUGGING = False Then
    
            If Not VersionOK(Version) Then
                Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
                Call CloseSocket(UserIndex)
                Exit Sub
        
            End If
    
        #End If
    
        If EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MacAddress, HDserial, MD5) Then
            Call WritePersonajesDeCuenta(UserIndex)
            Call WriteMostrarCuenta(UserIndex)
        Else
            
            Call CloseSocket(UserIndex)
            Exit Sub
    
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBorrarPJ(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim UserDelete     As String
        Dim CuentaEmail    As String
        Dim CuentaPassword As String
        Dim MacAddress     As String
        Dim HDserial       As Long
        Dim MD5            As String
        Dim Version        As String
        
        UserDelete = .incomingData.ReadASCIIString()
        CuentaEmail = .incomingData.ReadASCIIString()
        CuentaPassword = .incomingData.ReadASCIIString()
        Version = CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte())
        MacAddress = .incomingData.ReadASCIIString()
        HDserial = .incomingData.ReadLong()
        MD5 = .incomingData.ReadASCIIString()
        
        #If DEBUGGING = False Then
    
            If Not VersionOK(Version) Then
                Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
                Call CloseSocket(UserIndex)
                Exit Sub
    
            End If
    
        #End If
        
        If Not EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MacAddress, HDserial, MD5) Then
            Call CloseSocket(UserIndex)
            Exit Sub
    
        End If
        
        If Not CheckUserAccount(UserDelete, UserList(UserIndex).AccountId) Then
            Call LogHackAttemp(CuentaEmail & "[" & UserList(UserIndex).ip & "] intentó borrar el pj " & UserDelete)
            Call CloseSocket(UserIndex)
            Exit Sub
    
        End If
        
        ' Si está online el personaje a borrar, lo kickeo para prevenir dupeos.
        Dim targetUserIndex As Integer
        targetUserIndex = NameIndex(UserDelete)
    
        If targetUserIndex > 0 Then
            Call LogHackAttemp("Se trató de eliminar al personaje " & UserDelete & " cuando este estaba conectado desde la IP " & UserList(UserIndex).ip)
            Call CloseSocket(targetUserIndex)
    
        End If
    
        Call BorrarUsuarioDatabase(UserDelete)
        Call WritePersonajesDeCuenta(UserIndex)
  
    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim Seconds As Byte
        
        Seconds = .incomingData.ReadByte()

        If Not .flags.Privilegios And PlayerType.user Then
            CuentaRegresivaTimer = Seconds
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("íEmpezando cuenta regresiva desde: " & Seconds & " segundos...!", FontTypeNames.FONTTYPE_GUILD))
        
            
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCuentaRegresiva", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandlePossUser(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim UserName As String
        
        UserName = .incomingData.ReadASCIIString()
            
        If NameIndex(UserName) <= 0 Then
        
            If Not .flags.Privilegios And PlayerType.user Then
            
                If Database_Enabled Then
                    If Not SetPositionDatabase(UserName, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
                        Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)

                End If
                    
                If Not .flags.Privilegios And PlayerType.Consejero Then
                    Call WriteConsoleMsg(UserIndex, "Servidor> Acción realizada con exito! La nueva posicion de " & UserName & "es: " & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y & "...", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Servidor> Acción realizada con exito!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> El usuario debe estar deslogueado para dicha solicitud!", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePossUser", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleDuel(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
        
    Dim Players         As String
    Dim Bet             As Long
    Dim PocionesMaximas As Integer
    Dim CaenItems       As Boolean

    With UserList(UserIndex)

        Players = .incomingData.ReadASCIIString
        Bet = .incomingData.ReadLong
        PocionesMaximas = .incomingData.ReadInteger
        CaenItems = .incomingData.ReadBoolean

        Call CrearReto(UserIndex, Players, Bet, PocionesMaximas, CaenItems)

    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDuel", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleAcceptDuel(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
        
    Dim Offerer As String

    With UserList(UserIndex)

        Offerer = .incomingData.ReadASCIIString

        Call AceptarReto(UserIndex, Offerer)

    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAcceptDuel", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCancelDuel(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        .incomingData.ReadInteger

        If .flags.SolicitudReto.estado <> SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")

        ElseIf .flags.AceptoReto > 0 Then
            Call CancelarSolicitudReto(.flags.AceptoReto, .name & " ha cancelado su admisión.")

        End If

    End With

End Sub

Private Sub HandleQuitDuel(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .flags.EnReto Then
            Call AbandonarReto(UserIndex)
        End If

    End With

End Sub

Private Sub HandleNieveToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleNieveToggle_Err

    'Author: Pablo Mercavides
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        Call LogGM(.name, "/NIEVE")
        
        Nebando = Not Nebando
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

    End With
        
    Exit Sub

HandleNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleNieblaToggle(ByVal UserIndex As Integer)
        
    On Error GoTo HandleNieblaToggle_Err

    'Author: Pablo Mercavides
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.user Or PlayerType.Consejero)) Then Exit Sub
        
        Call LogGM(.name, "/NIEBLA")
        Call ResetMeteo

    End With
        
    Exit Sub

HandleNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleTransFerGold(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Cantidad As Long
        Dim tUser    As Integer
        
        Cantidad = .incomingData.ReadLong()
        UserName = .incomingData.ReadASCIIString()

        ' WyroX: Chequeos de seguridad... Estos chequeos ya se hacen en el cliente, pero si no se hacen se puede duplicar oro...

        ' Cantidad válida?
        If Cantidad <= 0 Then Exit Sub

        ' Tiene el oro?
        If .Stats.Banco < Cantidad Then Exit Sub
            
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
            
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        tUser = NameIndex(UserName)

        ' Enviar a vos mismo?
        If tUser = UserIndex Then
            Call WriteChatOverHead(UserIndex, "¡No puedo enviarte oro a vos mismo!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If
    
        If Not EsGM(UserIndex) Then

            If tUser <= 0 Then
                If Database_Enabled Then
                    If Not AddOroBancoDatabase(UserName, Cantidad) Then
                        Call WriteChatOverHead(UserIndex, "El usuario no existe.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If

                Else
                    Dim FileUser  As String
                    Dim OroenBove As Long
    
                    FileUser = CharPath & UCase$(UserName) & ".chr"
                    OroenBove = val(GetVar(FileUser, "STATS", "BANCO"))
                    OroenBove = OroenBove + val(Cantidad)
    
                    Call WriteVar(FileUser, "STATS", "BANCO", CLng(OroenBove)) 'Guardamos en bove

                End If

            Else
                UserList(tUser).Stats.Banco = UserList(tUser).Stats.Banco + val(Cantidad) 'Se lo damos al otro.

            End If
                
            UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
    
            Call WriteChatOverHead(UserIndex, "¡El envío se ha realizado con éxito! Gracias por utilizar los servicios de Finanzas Goliath", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("173", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
        Else
            Call WriteChatOverHead(UserIndex, "Los administradores no pueden transferir oro.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Call LogGM(.name, "Quizo transferirle oro a: " & UserName)
            
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim SlotViejo As Byte
        Dim SlotNuevo As Byte
        
        SlotViejo = .incomingData.ReadByte()
        SlotNuevo = .incomingData.ReadByte()
        
        Dim Objeto    As obj
        Dim Equipado  As Boolean
        Dim Equipado2 As Boolean
        Dim Equipado3 As Boolean
        
        If (SlotViejo > .CurrentInventorySlots) Or (SlotNuevo > .CurrentInventorySlots) Then
            Call WriteConsoleMsg(UserIndex, "Espacio no desbloqueado.", FontTypeNames.FONTTYPE_INFOIAO)
            
        Else
    
            If .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex Then
                .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotNuevo).amount + .Invent.Object(SlotViejo).amount
                    
                Dim Excedente As Integer
                Excedente = .Invent.Object(SlotNuevo).amount - MAX_INVENTORY_OBJS

                If Excedente > 0 Then
                    .Invent.Object(SlotViejo).amount = Excedente
                    .Invent.Object(SlotNuevo).amount = MAX_INVENTORY_OBJS
                Else

                    If .Invent.Object(SlotViejo).Equipped = 1 Then
                        .Invent.Object(SlotNuevo).Equipped = 1

                    End If
                    
                    .Invent.Object(SlotViejo).ObjIndex = 0
                    .Invent.Object(SlotViejo).amount = 0
                    .Invent.Object(SlotViejo).Equipped = 0
                    
                    'Cambiamos si alguno es un anillo
                    If .Invent.DañoMagicoEqpSlot = SlotViejo Then
                        .Invent.DañoMagicoEqpSlot = SlotNuevo

                    End If

                    If .Invent.ResistenciaEqpSlot = SlotViejo Then
                        .Invent.ResistenciaEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un armor
                    If .Invent.ArmourEqpSlot = SlotViejo Then
                        .Invent.ArmourEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un barco
                    If .Invent.BarcoSlot = SlotViejo Then
                        .Invent.BarcoSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es una montura
                    If .Invent.MonturaSlot = SlotViejo Then
                        .Invent.MonturaSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un casco
                    If .Invent.CascoEqpSlot = SlotViejo Then
                        .Invent.CascoEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un escudo
                    If .Invent.EscudoEqpSlot = SlotViejo Then
                        .Invent.EscudoEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es municiín
                    If .Invent.MunicionEqpSlot = SlotViejo Then
                        .Invent.MunicionEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un arma
                    If .Invent.WeaponEqpSlot = SlotViejo Then
                        .Invent.WeaponEqpSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un nudillo
                    If .Invent.NudilloSlot = SlotViejo Then
                        .Invent.NudilloSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es un magico
                    If .Invent.MagicoSlot = SlotViejo Then
                        .Invent.MagicoSlot = SlotNuevo

                    End If
                        
                    'Cambiamos si alguno es una herramienta
                    If .Invent.HerramientaEqpSlot = SlotViejo Then
                        .Invent.HerramientaEqpSlot = SlotNuevo

                    End If

                End If
                
            Else

                If .Invent.Object(SlotNuevo).ObjIndex <> 0 Then
                    Objeto.amount = .Invent.Object(SlotViejo).amount
                    Objeto.ObjIndex = .Invent.Object(SlotViejo).ObjIndex
                    
                    If .Invent.Object(SlotViejo).Equipped = 1 Then
                        Equipado = True
    
                    End If
                    
                    If .Invent.Object(SlotNuevo).Equipped = 1 Then
                        Equipado2 = True
    
                    End If
                    
                    '  If .Invent.Object(SlotNuevo).Equipped = 1 And .Invent.Object(SlotViejo).Equipped = 1 Then
                    '     Equipado3 = True
                    ' End If
                    
                    .Invent.Object(SlotViejo).ObjIndex = .Invent.Object(SlotNuevo).ObjIndex
                    .Invent.Object(SlotViejo).amount = .Invent.Object(SlotNuevo).amount
                    
                    .Invent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
                    .Invent.Object(SlotNuevo).amount = Objeto.amount
                    
                    If Equipado Then
                        .Invent.Object(SlotNuevo).Equipped = 1
                    Else
                        .Invent.Object(SlotNuevo).Equipped = 0
    
                    End If
                                    
                    If Equipado2 Then
                        .Invent.Object(SlotViejo).Equipped = 1
                    Else
                        .Invent.Object(SlotViejo).Equipped = 0
    
                    End If
    
                End If
    
                'Cambiamos si alguno es un anillo
                If .Invent.DañoMagicoEqpSlot = SlotViejo Then
                    .Invent.DañoMagicoEqpSlot = SlotNuevo
                ElseIf .Invent.DañoMagicoEqpSlot = SlotNuevo Then
                    .Invent.DañoMagicoEqpSlot = SlotViejo

                End If

                If .Invent.ResistenciaEqpSlot = SlotViejo Then
                    .Invent.ResistenciaEqpSlot = SlotNuevo
                ElseIf .Invent.ResistenciaEqpSlot = SlotNuevo Then
                    .Invent.ResistenciaEqpSlot = SlotViejo

                End If
                    
                'Cambiamos si alguno es un armor
                If .Invent.ArmourEqpSlot = SlotViejo Then
                    .Invent.ArmourEqpSlot = SlotNuevo
                ElseIf .Invent.ArmourEqpSlot = SlotNuevo Then
                    .Invent.ArmourEqpSlot = SlotViejo
    
                End If
                    
                'Cambiamos si alguno es un barco
                If .Invent.BarcoSlot = SlotViejo Then
                    .Invent.BarcoSlot = SlotNuevo
                ElseIf .Invent.BarcoSlot = SlotNuevo Then
                    .Invent.BarcoSlot = SlotViejo
    
                End If
                     
                'Cambiamos si alguno es una montura
                If .Invent.MonturaSlot = SlotViejo Then
                    .Invent.MonturaSlot = SlotNuevo
                ElseIf .Invent.MonturaSlot = SlotNuevo Then
                    .Invent.MonturaSlot = SlotViejo
    
                End If
                    
                'Cambiamos si alguno es un casco
                If .Invent.CascoEqpSlot = SlotViejo Then
                    .Invent.CascoEqpSlot = SlotNuevo
                ElseIf .Invent.CascoEqpSlot = SlotNuevo Then
                    .Invent.CascoEqpSlot = SlotViejo
    
                End If
                    
                'Cambiamos si alguno es un escudo
                If .Invent.EscudoEqpSlot = SlotViejo Then
                    .Invent.EscudoEqpSlot = SlotNuevo
                ElseIf .Invent.EscudoEqpSlot = SlotNuevo Then
                    .Invent.EscudoEqpSlot = SlotViejo
    
                End If
                    
                'Cambiamos si alguno es municiín
                If .Invent.MunicionEqpSlot = SlotViejo Then
                    .Invent.MunicionEqpSlot = SlotNuevo
                ElseIf .Invent.MunicionEqpSlot = SlotNuevo Then
                    .Invent.MunicionEqpSlot = SlotViejo
    
                End If
                    
                'Cambiamos si alguno es un arma
                If .Invent.WeaponEqpSlot = SlotViejo Then
                    .Invent.WeaponEqpSlot = SlotNuevo
                ElseIf .Invent.WeaponEqpSlot = SlotNuevo Then
                    .Invent.WeaponEqpSlot = SlotViejo
    
                End If
                     
                'Cambiamos si alguno es un nudillo
                If .Invent.NudilloSlot = SlotViejo Then
                    .Invent.NudilloSlot = SlotNuevo
                ElseIf .Invent.NudilloSlot = SlotNuevo Then
                    .Invent.NudilloSlot = SlotViejo
    
                End If
                     
                'Cambiamos si alguno es un magico
                If .Invent.MagicoSlot = SlotViejo Then
                    .Invent.MagicoSlot = SlotNuevo
                ElseIf .Invent.MagicoSlot = SlotNuevo Then
                    .Invent.MagicoSlot = SlotViejo
    
                End If
                     
                'Cambiamos si alguno es una herramienta
                If .Invent.HerramientaEqpSlot = SlotViejo Then
                    .Invent.HerramientaEqpSlot = SlotNuevo
                ElseIf .Invent.HerramientaEqpSlot = SlotNuevo Then
                    .Invent.HerramientaEqpSlot = SlotViejo
    
                End If
                
                If Objeto.ObjIndex = 0 Then
                    .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex
                    .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotViejo).amount
                    .Invent.Object(SlotNuevo).Equipped = .Invent.Object(SlotViejo).Equipped
                            
                    .Invent.Object(SlotViejo).ObjIndex = 0
                    .Invent.Object(SlotViejo).amount = 0
                    .Invent.Object(SlotViejo).Equipped = 0
    
                End If
                    
            End If
                
            Call UpdateUserInv(False, UserIndex, SlotViejo)
            Call UpdateUserInv(False, UserIndex, SlotNuevo)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMoveItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBovedaMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim SlotViejo As Byte
        Dim SlotNuevo As Byte
        
        SlotViejo = .incomingData.ReadByte()
        SlotNuevo = .incomingData.ReadByte()
        
        Dim Objeto    As obj
        Dim Equipado  As Boolean
        Dim Equipado2 As Boolean
        Dim Equipado3 As Boolean
        
        Objeto.ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex
        Objeto.amount = UserList(UserIndex).BancoInvent.Object(SlotViejo).amount
        
        UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotViejo).amount = UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount
         
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount = Objeto.amount
    
        'Actualizamos el banco
        Call UpdateBanUserInv(False, UserIndex, SlotViejo)
        Call UpdateBanUserInv(False, UserIndex, SlotNuevo)

    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBovedaMoveItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleQuieroFundarClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then Exit Sub

        If UserList(UserIndex).GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya perteneces a un clan, no podés fundar otro.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        If UserList(UserIndex).Stats.ELV < 35 Or UserList(UserIndex).Stats.UserSkills(eSkill.liderazgo) < 100 Then
            Call WriteConsoleMsg(UserIndex, "Para fundar un clan debes ser nivel 35, tener 100 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1).", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        If Not TieneObjetos(407, 1, UserIndex) Or Not TieneObjetos(408, 1, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Para fundar un clan debes tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1).", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        Call WriteConsoleMsg(UserIndex, "Servidor> ¡Comenzamos a fundar el clan! Ingresa todos los datos solicitados.", FontTypeNames.FONTTYPE_INFOIAO)
        
        Call WriteShowFundarClanForm(UserIndex)

    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuieroFundarClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleLlamadadeClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim refError   As String
        Dim clan_nivel As Byte
                        
        If .GuildIndex <> 0 Then
            clan_nivel = modGuilds.NivelDeClan(.GuildIndex)

            If clan_nivel > 1 Then
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> [" & .name & "] solicita apoyo de su clan en " & DarNameMapa(.Pos.Map) & " (" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & "). Puedes ver su ubicación en el mapa del mundo.", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.Pos.Map, .Pos.X, .Pos.Y))
            
            Else
                Call WriteConsoleMsg(UserIndex, "Servidor> El nivel de tu clan debe ser 2 para utilizar esta opción.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        Else
        
            Call WriteConsoleMsg(UserIndex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLlamadadeClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub


Private Sub HandleGenio(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleGenio_Err

    With UserList(UserIndex)

        'Si no es GM, no pasara nada :P
        If (.flags.Privilegios And PlayerType.user) Then Exit Sub
        
        Dim i As Byte

        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 100
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Tus skills fueron editados.", FontTypeNames.FONTTYPE_INFOIAO)

    End With
        
    Exit Sub

HandleGenio_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGenio", Erl)
        
End Sub

Private Sub HandleCasamiento(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer

        UserName = .incomingData.ReadASCIIString()
        tUser = NameIndex(UserName)
            
        If .flags.TargetNPC > 0 Then

            If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Then
                Call WriteConsoleMsg(UserIndex, "Primero haz click sobre un sacerdote.", FontTypeNames.FONTTYPE_INFO)

            Else

                If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede casarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        
                Else
            
                    If tUser = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "No podés casarte contigo mismo.", FontTypeNames.FONTTYPE_INFO)
                        
                    ElseIf .flags.Casado = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡Ya estás casado! Debes divorciarte de tu actual pareja para casarte nuevamente.", FontTypeNames.FONTTYPE_INFO)
                            
                    ElseIf UserList(tUser).flags.Casado = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Tu pareja debe divorciarse antes de tomar tu mano en matrimonio.", FontTypeNames.FONTTYPE_INFO)
                            
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)

                        Else

                            If UserList(tUser).flags.Candidato = UserIndex Then

                                UserList(tUser).flags.Casado = 1
                                UserList(tUser).flags.Pareja = UserList(UserIndex).name
                                .flags.Casado = 1
                                .flags.Pareja = UserList(tUser).name

                                Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(FXSound.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
                                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sacerdote de " & DarNameMapa(.Pos.Map) & " celebra el casamiento entre " & UserList(UserIndex).name & " y " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_WARNING))
                                Call WriteChatOverHead(UserIndex, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteChatOverHead(tUser, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                                
                            Else
                                
                                Call WriteChatOverHead(UserIndex, "La solicitud de casamiento a sido enviada a " & UserName & ".", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteConsoleMsg(tUser, .name & " desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER " & .name & ".", FontTypeNames.FONTTYPE_TALK)

                                .flags.Candidato = tUser

                            End If

                        End If

                    End If

                End If

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click sobre el sacerdote.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCasamiento", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleEnviarCodigo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Codigo As String

        Codigo = .incomingData.ReadASCIIString()

        Call CheckearCodigo(UserIndex, Codigo)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEnviarCodigo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCrearTorneo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)
 
        Dim NivelMinimo As Byte

        Dim nivelmaximo As Byte

        Dim cupos       As Byte

        Dim costo       As Long

        Dim mago        As Byte

        Dim clerico     As Byte

        Dim guerrero    As Byte

        Dim asesino     As Byte

        Dim bardo       As Byte

        Dim druido      As Byte

        Dim Paladin     As Byte

        Dim cazador     As Byte

        Dim Trabajador  As Byte

        Dim Mapa        As Integer

        Dim X           As Byte

        Dim Y           As Byte

        Dim nombre      As String

        Dim reglas      As String

        NivelMinimo = .incomingData.ReadByte
        nivelmaximo = .incomingData.ReadByte
        cupos = .incomingData.ReadByte
        costo = .incomingData.ReadLong
        mago = .incomingData.ReadByte
        clerico = .incomingData.ReadByte
        guerrero = .incomingData.ReadByte
        asesino = .incomingData.ReadByte
        bardo = .incomingData.ReadByte
        druido = .incomingData.ReadByte
        Paladin = .incomingData.ReadByte
        cazador = .incomingData.ReadByte
 
        Trabajador = .incomingData.ReadByte

        Mapa = .incomingData.ReadInteger
        X = .incomingData.ReadByte
        Y = .incomingData.ReadByte
        nombre = .incomingData.ReadASCIIString
        reglas = .incomingData.ReadASCIIString
  
        If EsGM(UserIndex) Then
            Torneo.NivelMinimo = NivelMinimo
            Torneo.nivelmaximo = nivelmaximo
            Torneo.cupos = cupos
            Torneo.costo = costo
            Torneo.mago = mago
            Torneo.clerico = clerico
            Torneo.guerrero = guerrero
            Torneo.asesino = asesino
            Torneo.bardo = bardo
            Torneo.druido = druido
            Torneo.Paladin = Paladin
            Torneo.cazador = cazador
            Torneo.Trabajador = Trabajador
        
            Torneo.Mapa = Mapa
            Torneo.X = X
            Torneo.Y = Y
            Torneo.nombre = nombre
            Torneo.reglas = reglas

            Call IniciarTorneo

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCrearTorneo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleComenzarTorneo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If EsGM(UserIndex) Then

            Call ComenzarTorneoOk

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCancelarTorneo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If EsGM(UserIndex) Then
            Call ResetearTorneo

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Tipo As Byte
            Tipo = .incomingData.ReadByte()
  
        If (.flags.Privilegios And Not (PlayerType.Consejero Or PlayerType.user)) Then

            Select Case Tipo

                Case 0

                    If Not BusquedaTesoroActiva And BusquedaRegaloActiva = False And BusquedaNpcActiva = False Then
                        Call PerderTesoro
                    Else

                        If BusquedaTesoroActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & "). ¿Quien sera el valiente que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & TesoroNumMapa & "-" & TesoroX & "-" & TesoroY, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                Case 1

                    If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
                        Call PerderRegalo
                    Else

                        If BusquedaRegaloActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningún valiente fue capaz de encontrar el item misterioso, recuerda que se encuentra en " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & "). ¡Ten cuidado!", FontTypeNames.FONTTYPE_TALK))
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & RegaloNumMapa & "-" & RegaloX & "-" & RegaloY, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                Case 2

                    If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                        Dim Pos As WorldPos
                        Pos.Map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
                        Pos.Y = 50
                        Pos.X = 50
                        npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), Pos, True, False, True)
                        BusquedaNpcActiva = True
                    Else

                        If BusquedaNpcActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavía nadie logró matar el NPC que se encuentra en el mapa " & NpcList(npc_index_evento).Pos.Map & ".", FontTypeNames.FONTTYPE_TALK))
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda de npc activo. El tesoro se encuentra en: " & NpcList(npc_index_evento).Pos.Map & "-" & NpcList(npc_index_evento).Pos.X & "-" & NpcList(npc_index_evento).Pos.Y, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

            End Select

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleFlagTrabajar(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        .Counters.Trabajando = 0
        .flags.UsandoMacro = False
        .flags.TargetObj = 0 ' Sacamos el targer del objeto
        .flags.UltimoMensaje = 0

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleEscribiendo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If .flags.Escribiendo = False Then
            .flags.Escribiendo = True
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetEscribiendo(.Char.CharIndex, True))
            
        Else
            .flags.Escribiendo = False
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetEscribiendo(.Char.CharIndex, False))

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleRequestFamiliar(ByVal UserIndex As Integer)
 
    On Error GoTo HandleRequestFamiliar_Err

    Call WriteFamiliar(UserIndex)
        
    Exit Sub

HandleRequestFamiliar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestFamiliar", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleCompletarAccion(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Accion As Byte
            Accion = .incomingData.ReadByte()
        
        If .Accion.AccionPendiente = True Then
            If .Accion.TipoAccion = Accion Then
                Call CompletarAccionFin(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Servidor> La acción que solicitas no se corresponde.", FontTypeNames.FONTTYPE_SERVER)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> Tu no tenias ninguna acción pendiente. ", FontTypeNames.FONTTYPE_SERVER)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleReclamarRecompensa(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim index As Byte
            index = .incomingData.ReadByte()
        
        Call EntregarRecompensas(UserIndex, index)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleTraerRecompensas(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Call EnviarRecompensaStat(UserIndex)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCorreo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        'Call WriteListaCorreo(Userindex, False)
        'Call EnviarRecompensaStat(UserIndex)

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCorreo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleSendCorreo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Nick               As String
        Dim msg                As String
        Dim ItemCount          As Byte
        Dim cant               As Integer
        Dim IndexReceptor      As Integer
        Dim Itemlista(1 To 10) As obj

        Nick = .incomingData.ReadASCIIString()
        msg = .incomingData.ReadASCIIString()
        ItemCount = .incomingData.ReadByte()
        
        Dim ObjIndex   As Integer
        Dim FinalCount As Byte
        Dim HuboError  As Boolean
                
        If ItemCount > 0 Then 'Si el correo tiene item

            Dim i As Byte

            For i = 1 To ItemCount
                Itemlista(i).ObjIndex = .incomingData.ReadByte
                Itemlista(i).amount = .incomingData.ReadInteger
            Next i

        Else 'Si es solo texto
            'IndexReceptor = NameIndex(Nick)
            FinalCount = 0
            AddCorreo UserIndex, Nick, msg, 0, FinalCount

        End If
        
        Dim ObjArray As String
        
        ' WyroX: Deshabilitado
        If False Then

            For i = 1 To ItemCount
                ObjIndex = UserList(UserIndex).Invent.Object(Itemlista(i).ObjIndex).ObjIndex
                
                If ObjData(ObjIndex).Destruye = 1 Then
                    HuboError = True
                Else

                    If ObjData(ObjIndex).Instransferible = 1 Then
                        HuboError = True
                        '  Call WriteConsoleMsg(UserIndex, "No podes transferir ese item.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If ObjData(ObjIndex).Newbie = 1 Then
                            HuboError = True
                            ' Call WriteConsoleMsg(UserIndex, "No podes transferir ese item.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If ObjData(ObjIndex).Intirable = 1 Then
                                HuboError = True
                                ' Call WriteConsoleMsg(UserIndex, "No podes transferir ese item.", FontTypeNames.FONTTYPE_INFO)
                            Else

                                If ObjData(ObjIndex).OBJType = eOBJType.otMonturas And UserList(UserIndex).flags.Montado Then
                                    HuboError = True
                                    '  Call WriteConsoleMsg(UserIndex, "Para transferir tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
                                Else
                                
                                    Call QuitarUserInvItem(UserIndex, Itemlista(i).ObjIndex, Itemlista(i).amount)
                                    Call UpdateUserInv(False, UserIndex, Itemlista(i).ObjIndex)
                                    FinalCount = FinalCount + 1
                                    ObjArray = ObjArray & ObjIndex & "-" & Itemlista(i).amount & "@"

                                End If

                            End If

                        End If

                    End If

                End If

            Next i
                
            IndexReceptor = NameIndex(Nick)
            AddCorreo UserIndex, Nick, msg, ObjArray, FinalCount
    
            If HuboError Then
                Call WriteConsoleMsg(UserIndex, "Hubo objetos que no se pudieron enviar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
        Else
            Call WriteConsoleMsg(UserIndex, "Correo desactivado.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendCorreo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleRetirarItemCorreo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim MsgIndex As Integer

        MsgIndex = .incomingData.ReadInteger()
        
        'Call ExtractItemCorreo(Userindex, MsgIndex)

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRetirarItemCorreo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBorrarCorreo(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim MsgIndex As Integer

        MsgIndex = .incomingData.ReadInteger()
        
        'Call BorrarCorreoMail(Userindex, MsgIndex)

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBorrarCorreo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleInvitarGrupo(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            
        Else
            
            If .Grupo.CantidadMiembros <= UBound(.Grupo.Miembros) Then
                Call WriteWorkRequestTarget(UserIndex, eSkill.Grupo)
            Else
                Call WriteConsoleMsg(UserIndex, "¡No podés invitar a más personas!", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With
        
    Exit Sub

HandleInvitarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInvitarGrupo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleMarcaDeClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleMarcaDeClan_Err

    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
       
        Call WriteWorkRequestTarget(UserIndex, eSkill.MarcaDeClan)

    End With
        
    Exit Sub

HandleMarcaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMarcaDeClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
End Sub

Private Sub HandleMarcaDeGM(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleMarcaDeGM_Err

    Call WriteWorkRequestTarget(UserIndex, eSkill.MarcaDeGM)

    Exit Sub

HandleMarcaDeGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMarcaDeGM", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleResponderPregunta(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim respuesta As Boolean
        Dim DeDonde   As String

        respuesta = .incomingData.ReadBoolean()
        
        Dim Log As String

        Log = "Repuesta "

        If respuesta Then
        
            Select Case UserList(UserIndex).flags.pregunta

                Case 1
                    Log = "Repuesta Afirmativa 1"

                    'Call WriteConsoleMsg(UserIndex, "El usuario desea unirse al grupo.", FontTypeNames.FONTTYPE_SUBASTA)
                    ' UserList(UserIndex).Grupo.PropuestaDe = 0
                    If UserList(UserIndex).Grupo.PropuestaDe <> 0 Then
                
                        If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Lider <> UserList(UserIndex).Grupo.PropuestaDe Then
                            Call WriteConsoleMsg(UserIndex, "íEl lider del grupo a cambiado, imposible unirse!", FontTypeNames.FONTTYPE_INFOIAO)
                        Else
                        
                            Log = "Repuesta Afirmativa 1-1 "
                        
                            If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Lider = 0 Then
                                Call WriteConsoleMsg(UserIndex, "íEl grupo ya no existe!", FontTypeNames.FONTTYPE_INFOIAO)
                            Else
                            
                                Log = "Repuesta Afirmativa 1-2 "
                            
                                If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros = 1 Then
                                    Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe, "36", FontTypeNames.FONTTYPE_INFOIAO)
                                    'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "íEl grupo a sido creado!", FontTypeNames.FONTTYPE_INFOIAO)
                                    UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.EnGrupo = True
                                    Log = "Repuesta Afirmativa 1-3 "

                                End If
                                
                                Log = "Repuesta Afirmativa 1-4"
                                UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros = UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros + 1
                                UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Miembros(UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros) = UserIndex
                                UserList(UserIndex).Grupo.EnGrupo = True
                                
                                Dim index As Byte
                                
                                Log = "Repuesta Afirmativa 1-5 "
                                
                                For index = 2 To UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros - 1
                                    Call WriteLocaleMsg(UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Miembros(index), "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
                                
                                Next index
                                
                                Log = "Repuesta Afirmativa 1-6 "
                                'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "í" & UserList(UserIndex).name & " a sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe, "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
                                
                                Call WriteConsoleMsg(UserIndex, "¡Has sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                
                                Log = "Repuesta Afirmativa 1-7 "
                                
                                Call RefreshCharStatus(UserList(UserIndex).Grupo.PropuestaDe)
                                Call RefreshCharStatus(UserIndex)
                                 
                                Log = "Repuesta Afirmativa 1-8"

                                Call CompartirUbicacion(UserIndex)

                            End If

                        End If

                    Else
                    
                        Call WriteConsoleMsg(UserIndex, "Servidor> Solicitud de grupo invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                    
                    End If

                    'unirlo
                Case 2
                    Log = "Repuesta Afirmativa 2"
                    Call WriteConsoleMsg(UserIndex, "¡Ahora sos un ciudadano!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call VolverCiudadano(UserIndex)
                    
                Case 3
                    Log = "Repuesta Afirmativa 3"
                    
                    UserList(UserIndex).Hogar = UserList(UserIndex).PosibleHogar

                    Select Case UserList(UserIndex).Hogar

                        Case eCiudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                            
                        Case eCiudad.cNix
                            DeDonde = "Nix"
                
                        Case eCiudad.cBanderbill
                            DeDonde = "Banderbill"
                        
                        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                            DeDonde = "Lindos"
                            
                        Case eCiudad.cArghal
                            DeDonde = " Arghal"
                            
                        Case eCiudad.cArkhein
                            DeDonde = " Arkhein"
                            
                        Case Else
                            DeDonde = "Ullathorpe"

                    End Select
                    
                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                    
                        Call WriteChatOverHead(UserIndex, "¡Gracias " & UserList(UserIndex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡Gracias " & UserList(UserIndex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                    
                Case 4
                    Log = "Repuesta Afirmativa 4"
                
                    If UserList(UserIndex).flags.TargetUser <> 0 Then
                
                        UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                        UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                        UserList(UserIndex).ComUsu.cant = 0
                        UserList(UserIndex).ComUsu.Objeto = 0
                        UserList(UserIndex).ComUsu.Acepto = False
                    
                        'Rutina para comerciar con otro usuario
                        Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)

                    Else
                        Call WriteConsoleMsg(UserIndex, "Servidor> Solicitud de comercio invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                
                    End If
                
                Case 5
                    Log = "Repuesta Afirmativa 5"
                
                    If MapInfo(.Pos.Map).Newbie Then
                        Call WarpToLegalPos(UserIndex, 140, 53, 58)
                        .Counters.TimerBarra = 5
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Resucitar, .Counters.TimerBarra, False))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, .Counters.TimerBarra, Accion_Barra.Resucitar))
                        UserList(UserIndex).Accion.AccionPendiente = True
                        UserList(UserIndex).Accion.Particula = ParticulasIndex.Resucitar
                        UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Resucitar
    
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", .Pos.X, .Pos.Y))
                        'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "82", FontTypeNames.FONTTYPE_INFOIAO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ya no te encuentras en un mapa newbie.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                
                Case Else
                    Call WriteConsoleMsg(UserIndex, "No tienes preguntas pendientes.", FontTypeNames.FONTTYPE_INFOIAO)
                    
            End Select
        
        Else
            Log = "Repuesta negativa"
        
            Select Case UserList(UserIndex).flags.pregunta

                Case 1
                    Log = "Repuesta negativa 1"

                    If UserList(UserIndex).Grupo.PropuestaDe <> 0 Then
                        Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "El usuario no esta interesado en formar parte del grupo.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                    UserList(UserIndex).Grupo.PropuestaDe = 0
                    Call WriteConsoleMsg(UserIndex, "Has rechazado la propuesta.", FontTypeNames.FONTTYPE_INFOIAO)
                
                Case 2
                    Log = "Repuesta negativa 2"
                    Call WriteConsoleMsg(UserIndex, "¡Continuas siendo neutral!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call VolverCriminal(UserIndex)

                Case 3
                    Log = "Repuesta negativa 3"
                    
                    Select Case UserList(UserIndex).PosibleHogar

                        Case eCiudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                            
                        Case eCiudad.cNix
                            DeDonde = "Nix"
                
                        Case eCiudad.cBanderbill
                            DeDonde = "Banderbill"
                        
                        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                            DeDonde = "Lindos"
                            
                        Case eCiudad.cArghal
                            DeDonde = " Arghal"
                            
                        Case eCiudad.cArkhein
                            DeDonde = " Arkhein"
                            
                        Case Else
                            DeDonde = "Ullathorpe"

                    End Select
                    
                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                        Call WriteChatOverHead(UserIndex, "¡No hay problema " & UserList(UserIndex).name & "! Sos bienvenido en " & DeDonde & " cuando gustes.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If

                    UserList(UserIndex).PosibleHogar = UserList(UserIndex).Hogar
                    
                Case 4
                    Log = "Repuesta negativa 4"
                    
                    If UserList(UserIndex).flags.TargetUser <> 0 Then
                        Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "El usuario no desea comerciar en este momento.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Case 5
                    Log = "Repuesta negativa 5"
                    'No hago nada. dijo que no lo resucite
                        
                Case Else
                    Call WriteConsoleMsg(UserIndex, "No tienes preguntas pendientes.", FontTypeNames.FONTTYPE_INFOIAO)

            End Select
            
        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleResponderPregunta", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleRequestGrupo(ByVal UserIndex As Integer)

    On Error GoTo hErr

    'Author: Pablo Mercavides

    Call WriteDatosGrupo(UserIndex)
    
    Exit Sub
    
hErr:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestGrupo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleAbandonarGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleAbandonarGrupo_Err

    With UserList(UserIndex)

        
        Call .incomingData.ReadInteger
        
        If UserList(UserIndex).Grupo.Lider = UserIndex Then
            
            Call FinalizarGrupo(UserIndex)

            Dim i As Byte
            
            For i = 2 To UserList(UserIndex).Grupo.CantidadMiembros
                Call WriteUbicacion(UserIndex, i, 0)
            Next i

            UserList(UserIndex).Grupo.CantidadMiembros = 0
            UserList(UserIndex).Grupo.EnGrupo = False
            UserList(UserIndex).Grupo.Lider = 0
            UserList(UserIndex).Grupo.PropuestaDe = 0
            Call WriteConsoleMsg(UserIndex, "Has disuelto el grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            Call RefreshCharStatus(UserIndex)
            
        Else
            Call SalirDeGrupo(UserIndex)

        End If

    End With
        
    Exit Sub

HandleAbandonarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAbandonarGrupo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleHecharDeGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleHecharDeGrupo_Err

    With UserList(UserIndex)

        Dim Indice As Byte

        Indice = .incomingData.ReadByte()
        
        Call EcharMiembro(UserIndex, Indice)

    End With
        
    Exit Sub

HandleHecharDeGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHecharDeGrupo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleMacroPos(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleMacroPos_Err

    With UserList(UserIndex)

        .ChatCombate = .incomingData.ReadByte()
        .ChatGlobal = .incomingData.ReadByte()

    End With
        
    Exit Sub

HandleMacroPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMacroPos", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleSubastaInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleSubastaInfo_Err

    With UserList(UserIndex)

        If Subasta.HaySubastaActiva Then

            Call WriteConsoleMsg(UserIndex, "Subastador: " & Subasta.Subastador, FontTypeNames.FONTTYPE_SUBASTA)
            Call WriteConsoleMsg(UserIndex, "Objeto: " & ObjData(Subasta.ObjSubastado).name & " (" & Subasta.ObjSubastadoCantidad & ")", FontTypeNames.FONTTYPE_SUBASTA)

            If Subasta.HuboOferta Then
                Call WriteConsoleMsg(UserIndex, "Mejor oferta: " & PonerPuntos(Subasta.MejorOferta) & " monedas de oro por " & Subasta.Comprador & ".", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & PonerPuntos(Subasta.MejorOferta + 100), FontTypeNames.FONTTYPE_SUBASTA)
            Else
                Call WriteConsoleMsg(UserIndex, "Oferta inicial: " & PonerPuntos(Subasta.OfertaInicial) & " monedas de oro.", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & PonerPuntos(Subasta.OfertaInicial + 100), FontTypeNames.FONTTYPE_SUBASTA)

            End If

            Call WriteConsoleMsg(UserIndex, "Tiempo Restante de subasta:  " & SumarTiempo(Subasta.TiempoRestanteSubasta), FontTypeNames.FONTTYPE_SUBASTA)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta activa en este momento.", FontTypeNames.FONTTYPE_SUBASTA)

        End If

    End With
        
    Exit Sub

HandleSubastaInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSubastaInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
End Sub

Private Sub HandleScrollInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
 
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim activo As Boolean
        Dim HR     As Integer
        Dim MS     As Integer
        Dim SS     As Integer
        Dim secs   As Integer
        
        If UserList(UserIndex).flags.ScrollExp > 1 Then
            secs = UserList(UserIndex).Counters.ScrollExperiencia
            HR = secs \ 3600
            MS = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                Call WriteConsoleMsg(UserIndex, "Scroll de experiencia activo. Tiempo restante: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(UserIndex, "Scroll de experiencia activo. Tiempo restante: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)

            End If

            activo = True

        End If

        If UserList(UserIndex).flags.ScrollOro > 1 Then
            secs = UserList(UserIndex).Counters.ScrollOro
            HR = secs \ 3600
            MS = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                Call WriteConsoleMsg(UserIndex, "Scroll de oro activo. Tiempo restante: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(UserIndex, "Scroll de oro activo. Tiempo restante: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)

            End If

            activo = True

        End If

        If Not activo Then
            Call WriteConsoleMsg(UserIndex, "No tenes ningun scroll activo.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleScrollInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCancelarExit(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleCancelarExit_Err

    Call CancelExit(UserIndex)
        
    Exit Sub

HandleCancelarExit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleBanCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Reason   As String
        
        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call BanAccount(UserIndex, UserName, Reason)
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBanCuenta", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleUnBanCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call UnBanAccount(UserIndex, UserName)
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUnBanCuenta", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
         
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanSerialOK(UserIndex, UserName)

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBanSerial", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleUnBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
         
        UserName = .incomingData.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call UnBanSerialOK(UserIndex, UserName)
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUnBanSerial", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCerrarCliente(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser    As Integer
         
        UserName = .incomingData.ReadASCIIString()
        
        ' Solo administradores pueden cerrar clientes ajenos
        If (.flags.Privilegios And PlayerType.Admin) Then

            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " cerro el cliente de " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    
                Call WriteCerrarleCliente(tUser)

                Call LogGM(.name, "Cerro el cliene de:" & UserName)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarCliente", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleEventoInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleEventoInfo_Err

    With UserList(UserIndex)

        If EventoActivo Then
            Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
        Else
            Call WriteConsoleMsg(UserIndex, "Eventos> Actualmente no hay ningun evento en curso.", FontTypeNames.FONTTYPE_New_Eventos)

        End If
        
        Dim i           As Byte
        Dim encontre    As Boolean
        Dim HoraProximo As Byte
   
        If Not HoraEvento + 1 >= 24 Then
   
            For i = HoraEvento + 1 To 23

                If Evento(i).Tipo <> 0 Then
                    encontre = True
                    HoraProximo = i
                    Exit For

                End If

            Next i

        End If
        
        If encontre = False Then

            For i = 0 To HoraEvento

                If Evento(i).Tipo <> 0 Then
                    encontre = True
                    HoraProximo = i
                    Exit For

                End If

            Next i

        End If
        
        If encontre Then
            Call WriteConsoleMsg(UserIndex, "Eventos> El proximo evento " & DescribirEvento(HoraProximo) & " iniciara a las " & HoraProximo & ":00 horas.", FontTypeNames.FONTTYPE_New_Eventos)
        Else
            Call WriteConsoleMsg(UserIndex, "Eventos> No hay eventos proximos.", FontTypeNames.FONTTYPE_New_Eventos)

        End If

    End With
        
    Exit Sub

HandleEventoInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
End Sub

Private Sub HandleCrearEvento(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Tipo           As Byte
        Dim Duracion       As Byte
        Dim multiplicacion As Byte
        
        Tipo = .incomingData.ReadByte()
        Duracion = .incomingData.ReadByte()
        multiplicacion = .incomingData.ReadByte()

        If multiplicacion > 3 Then 'no superar este multiplicador
            multiplicacion = 3
        End If
        
        '/ dejar solo Administradores
        If .flags.Privilegios >= PlayerType.Admin Then
            If EventoActivo = False Then
                If LenB(Tipo) = 0 Or LenB(Duracion) = 0 Or LenB(multiplicacion) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.", FontTypeNames.FONTTYPE_New_Eventos)
                Else
                
                    Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).name)
                  
                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.", FontTypeNames.FONTTYPE_New_Eventos)

            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleBanTemporal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
         
        Dim UserName As String
        Dim Reason   As String
        Dim dias     As Byte
        
        UserName = .incomingData.ReadASCIIString()
        Reason = .incomingData.ReadASCIIString()
        dias = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call Admin.BanTemporal(UserName, dias, Reason, UserList(UserIndex).name)
        End If

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.?", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleTraerShop(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleTraerShop_Err

    Call WriteShop(UserIndex)
        
    Exit Sub

HandleTraerShop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTraerShop", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
End Sub

Private Sub HandleTraerRanking(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
        
    On Error GoTo HandleTraerRanking_Err

    Call WriteRanking(UserIndex)
        
    Exit Sub

HandleTraerRanking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTraerRanking", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
End Sub

Private Sub HandleComprarItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim ItemIndex    As Byte
        Dim ObjComprado  As obj
        Dim LogeoDonador As String

        ItemIndex = .incomingData.ReadByte()
        
        Dim i              As Byte
        Dim InvSlotsLibres As Byte
        
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
        Next i
    
        'Nos fijamos si entra
        If InvSlotsLibres = 0 Then
            Call WriteConsoleMsg(UserIndex, "Donación> Sin espacio en el inventario.", FontTypeNames.FONTTYPE_WARNING)
        Else

            If CreditosDonadorCheck(UserList(UserIndex).Cuenta) - ObjDonador(ItemIndex).Valor >= 0 Then
                ObjComprado.amount = ObjDonador(ItemIndex).Cantidad
                ObjComprado.ObjIndex = ObjDonador(ItemIndex).ObjIndex
            
                LogeoDonador = LogeoDonador & vbCrLf & "****************************************************" & vbCrLf
                LogeoDonador = LogeoDonador & "Compra iniciada. Balance de la cuenta " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos." & vbCrLf
                LogeoDonador = LogeoDonador & "El personaje " & UserList(UserIndex).name & "(" & UserList(UserIndex).Cuenta & ") Compro el item " & ObjData(ObjDonador(ItemIndex).ObjIndex).name & vbCrLf
                LogeoDonador = LogeoDonador & "Se descontaron " & CLng(ObjDonador(ItemIndex).Valor) & " creditos de la cuenta " & UserList(UserIndex).Cuenta & "." & vbCrLf
            
                If Not MeterItemEnInventario(UserIndex, ObjComprado) Then
                    LogeoDonador = LogeoDonador & "El item se tiro al piso" & vbCrLf
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjComprado)

                End If
                
                LogeoDonador = LogeoDonador & "****************************************************" & vbCrLf
             
                Call RestarCreditosDonador(UserList(UserIndex).Cuenta, CLng(ObjDonador(ItemIndex).Valor))
                Call WriteConsoleMsg(UserIndex, "Donación> Gracias por tu compra. Tu saldo es de " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                Call LogearEventoDeDonador(LogeoDonador)
                Call SaveUser(UserIndex)
                Call WriteActShop(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Donación> Tu saldo es insuficiente. Actualmente tu saldo es de " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteActShop(UserIndex)

            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleComprarItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleCompletarViaje(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Destino As Byte

        Dim costo   As Long

        Destino = .incomingData.ReadByte()
        costo = .incomingData.ReadLong()

        ' WyroX: WTF el costo lo decide el cliente... Desactivo....
        Exit Sub

        If costo <= 0 Then Exit Sub

        Dim DeDonde As CityWorldPos

        If UserList(UserIndex).Stats.GLD < costo Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            
        Else

            Select Case Destino

                Case eCiudad.cUllathorpe
                    DeDonde = CityUllathorpe
                        
                Case eCiudad.cNix
                    DeDonde = CityNix
            
                Case eCiudad.cBanderbill
                    DeDonde = CityBanderbill
                    
                Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                    DeDonde = CityLindos
                        
                Case eCiudad.cArghal
                    DeDonde = CityArghal
                        
                Case eCiudad.cArkhein
                    DeDonde = CityArkhein
                        
                Case Else
                    DeDonde = CityUllathorpe

            End Select
        
            If DeDonde.NecesitaNave > 0 Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                    Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                        If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
                            Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                        End If

                    End If

                    Call WarpToLegalPos(UserIndex, DeDonde.MapaViaje, DeDonde.ViajeX, DeDonde.ViajeY, True)
                    Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
                    UserList(UserIndex).Stats.MinAGU = 0
                    UserList(UserIndex).Stats.MinHam = 0
                    UserList(UserIndex).flags.Sed = 1
                    UserList(UserIndex).flags.Hambre = 1
                    
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
                    Call WriteUpdateHungerAndThirst(UserIndex)
                    Call WriteUpdateUserStats(UserIndex)

                End If

            Else
            
                Dim Map As Integer

                Dim X   As Byte

                Dim Y   As Byte
            
                Map = DeDonde.MapaViaje
                X = DeDonde.ViajeX
                Y = DeDonde.ViajeY

                If UserList(UserIndex).flags.TargetNPC <> 0 Then
                    If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
                        Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                End If
                
                Call WarpUserChar(UserIndex, Map, X, Y, True)
                Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
                UserList(UserIndex).Stats.MinAGU = 0
                UserList(UserIndex).Stats.MinHam = 0
                UserList(UserIndex).flags.Sed = 1
                UserList(UserIndex).flags.Hambre = 1
                
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call WriteUpdateUserStats(UserIndex)
        
            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCompletarViaje", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
        
    On Error GoTo HandleQuest_Err

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete Quest.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex As Integer
    Dim tmpByte  As Byte

    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    'El NPC hace quests?
    If NpcList(NpcIndex).NumQuest = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", NpcList(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub

    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", NpcList(NpcIndex).Char.CharIndex, vbWhite))

    Exit Sub

HandleQuest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
        
    On Error GoTo HandleQuestAccept_Err

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el evento de aceptar una quest.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex  As Integer
    Dim QuestSlot As Byte
    Dim Indice    As Byte

    Indice = UserList(UserIndex).incomingData.ReadByte
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    If Indice = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
        
    If TieneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
        Call WriteConsoleMsg(UserIndex, "La quest ya esta en curso.", FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub

    End If
        
    'El personaje completo la quest que requiere?
    If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest > 0 Then
        If Not UserDoneQuest(UserIndex, QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest) Then
            Call WriteChatOverHead(UserIndex, "Debes completas la quest " & QuestList(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest).nombre & " para emprender esta mision.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub

        End If

    End If

    'El personaje tiene suficiente nivel?
    If UserList(UserIndex).Stats.ELV < QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel Then
        Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel & " para emprender esta mision.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If
        
    'El personaje ya hizo la quest?
    If UserDoneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
        Call WriteChatOverHead(UserIndex, "QUESTNEXT*" & NpcList(NpcIndex).QuestNumber(Indice), NpcList(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If
    
    QuestSlot = FreeQuestSlot(UserIndex)

    If QuestSlot = 0 Then
        Call WriteChatOverHead(UserIndex, "Debes completar las misiones en curso para poder aceptar más misiones.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If
    
    'Agregamos la quest.
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        .QuestIndex = NpcList(NpcIndex).QuestNumber(Indice)
        
        If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
        If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
        Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
        
    End With
        
    Exit Sub

HandleQuestAccept_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestAccept", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
        
    On Error GoTo HandleQuestDetailsRequest_Err

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestInfoRequest.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim QuestSlot As Byte

    QuestSlot = UserList(UserIndex).incomingData.ReadByte
    
    Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
        
    Exit Sub

HandleQuestDetailsRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestDetailsRequest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestAbandon.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    On Error GoTo HandleQuestAbandon_Err

    'Borramos la quest.
    Call CleanQuestSlot(UserIndex, UserList(UserIndex).incomingData.ReadByte)
    
    'Ordenamos la lista de quests del usuario.
    Call ArrangeUserQuests(UserIndex)
    
    'Enviamos la lista de quests actualizada.
    Call WriteQuestListSend(UserIndex)
        
    Exit Sub

HandleQuestAbandon_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestAbandon", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestListRequest.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
    On Error GoTo HandleQuestListRequest_Err

    Call WriteQuestListSend(UserIndex)
        
    Exit Sub

HandleQuestListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestListRequest", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo ErrHandler

    Dim Map   As Integer
    Dim X     As Byte
    Dim Y     As Byte
    Dim index As Long
    
    With UserList(UserIndex)

        Map = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        ' User Admin?
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posicion invalida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Choose pretorian clan index
        If Map = MAPA_PRETORIANO Then
            index = ePretorianType.Default ' Default clan
            
        Else
            index = ePretorianType.Custom ' Custom Clan

        End If
            
        ' Is already active any clan?
        If Not ClanPretoriano(index).Active Then
            
            If Not ClanPretoriano(index).SpawnClan(Map, X, Y, index) Then
                Call WriteConsoleMsg(UserIndex, "La posicion no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)

            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)

        End If
    
    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreatePretorianClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim Map   As Integer
    Dim index As Long
    
    With UserList(UserIndex)

        Map = .incomingData.ReadInteger()
        
        ' User Admin?
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa invalido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Los sacamos correctamente.
        Call EliminarPretorianos(Map)
    
    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreatePretorianClan", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

''
' Handles the "Consulta" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsulta(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/05/2010
    'Habilita/Deshabilita el modo consulta.
    '01/05/2010: ZaMa - Agrego validaciones.
    '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
    '***************************************************

    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
 
        Dim Nick As String
        Nick = .incomingData.ReadASCIIString

        ' Comando exclusivo para gms
        If Not EsGM(UserIndex) Then Exit Sub
        
        If Len(Nick) <> 0 Then
            UserConsulta = NameIndex(Nick)
            
            'Se asegura que el target exista
            If UserConsulta <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
            
        Else
        
            UserConsulta = .flags.TargetUser
            
            'Se asegura que el target exista
            If UserConsulta <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
            
        End If

        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGM(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserList(UserConsulta).name & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call LogGM(.name, "Termino consulta con " & UserList(UserConsulta).name)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
            ' Sino la inicia
        Else
        
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserList(UserConsulta).name & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call LogGM(.name, "Inicio consulta con " & UserList(UserConsulta).name)
            
            With UserList(UserConsulta)

                If Not EstaPCarea(UserIndex, UserConsulta) Then
                    Dim X As Byte
                    Dim Y As Byte
                        
                    X = .Pos.X
                    Y = .Pos.Y
                    Call FindLegalPos(UserIndex, .Pos.Map, X, Y)
                    Call WarpUserChar(UserIndex, .Pos.Map, X, Y, True)
                        
                End If
            
                If UserList(UserIndex).flags.AdminInvisible = 1 Then
                    Call DoAdminInvisible(UserIndex)

                End If

                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                            
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                    End If

                End If

            End With

        End If
        
        Call SetModoConsulta(UserConsulta)

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleConsulta", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleRequestProcesses(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim Nick As String
        Nick = .incomingData.ReadASCIIString

        ' Comando exclusivo para gms
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Dim tUser As Integer
    
            If Len(Nick) <> 0 Then
                tUser = NameIndex(Nick)
                
                'Se asegura que el target exista
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
            Else
            
                tUser = .flags.TargetUser
                
                'Se asegura que el target exista
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
    
            End If
    
            If tUser <> UserIndex Then
                If AdministratorAccounts.Exists(UCase$(UserList(tUser).name)) Then
                    Call WriteConsoleMsg(UserIndex, "No podés invadir la privacidad de otro administrador.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
            
            If LenB(UserList(tUser).flags.ProcesosPara) = 0 Then
                Call WriteRequestProcesses(tUser)

            End If
    
            UserList(tUser).flags.ProcesosPara = UserList(tUser).flags.ProcesosPara & ":" & .name
    
        End If

    End With

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestProcesses", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleRequestScreenShot(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim Nick As String
        Nick = .incomingData.ReadASCIIString

        ' Comando exclusivo para gms
            
        Dim tUser As Integer
            
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) = 0 Then Exit Sub

        If Len(Nick) <> 0 Then
            tUser = NameIndex(Nick)
            
            'Se asegura que el target exista
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
        Else
        
            tUser = .flags.TargetUser
            
            'Se asegura que el target exista
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        If tUser <> UserIndex Then
            If AdministratorAccounts.Exists(UCase$(UserList(tUser).name)) Then
                Call WriteConsoleMsg(UserIndex, "No podés invadir la privacidad de otro administrador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        If LenB(UserList(tUser).flags.ScreenShotPara) = 0 Then
            ' Creo un buffer de 2mb para la screenshot
            Set UserList(tUser).flags.ScreenShot = New clsByteQueue
            UserList(tUser).flags.ScreenShot.Capacity = 2097152
            
            Call WriteRequestScreenShot(tUser)

        End If

        UserList(tUser).flags.ScreenShotPara = UserList(tUser).flags.ScreenShotPara & ":" & .name

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRequestScreenShot", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleTolerancia0(ByVal UserIndex As Integer)


    With UserList(UserIndex)

        Dim Nick As String
        Nick = .incomingData.ReadASCIIString

        ' Comando exclusivo para admins
        If (.flags.Privilegios And PlayerType.Admin) = 0 Then Exit Sub
        
        Dim tUser As Integer

        tUser = NameIndex(Nick)
        
        'Se asegura que el target exista
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call WriteTolerancia0(tUser)
        Call BanIpAgrega(UserList(tUser).ip)
        Call BanSerialOK(UserIndex, Nick)
        Call BanAccount(UserIndex, Nick, "Tolerancia cero")

    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTolerancia0", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleScreenShot(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        On Error GoTo ErrHandler
        
        Dim data As String
        data = .incomingData.ReadASCIIString
           
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            
            ' Si nadie requirió esto, salimos
            If LenB(.flags.ScreenShotPara) = 0 Then Exit Sub
        
            Dim Finished As Boolean
        
            ' Por seguridad, limito a 10Kb de datos (dejo margen para el nombre y el resto del paquete)
            If LenB(data) = 0 Or Len(data) > 10000 Then
                data = "ERROR"
                Finished = True
        
                ' Si envió menos de 10Kb y termina con ~~~
            ElseIf Len(data) <= 10000 And Right$(data, 3) = "~~~" Then
                ' Damos la screenshot por terminada
                Finished = True

            End If

            ' Lo guardo en la cola
            Call .flags.ScreenShot.WriteASCIIStringFixed(data)
        
            If Finished Then
                Dim ListaGMs() As String
                ListaGMs = Split(.flags.ScreenShotPara, ":")
            
                Dim i As Integer, tGM As Integer, Offset As Long
    
                For i = LBound(ListaGMs) To UBound(ListaGMs)
                    tGM = NameIndex(ListaGMs(i))
                
                    If tGM > 0 Then
                    
                        For Offset = 0 To .flags.ScreenShot.Length - 1 Step 10000
                            Call WriteScreenShotData(tGM, .flags.ScreenShot, Offset, Min(.flags.ScreenShot.Length - Offset, 10000))
                        Next
                        
                        Call WriteShowScreenShot(tGM, .name)

                    End If

                Next

                .flags.ScreenShotPara = vbNullString
                Set .flags.ScreenShot = Nothing

            End If

        End If

    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleScreenShot", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleProcesses(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        On Error GoTo ErrHandler

        Dim data As String
        data = .incomingData.ReadASCIIString
        
        ' Por seguridad, limito a 10kb de datos (con margen para el nombre)
        If Len(data) > 10000 Then
            data = Left$(data, 10000) & vbNewLine & "[...Demasiado largo]"
        End If

        ' Si nadie requirió esto, salimos
        If LenB(.flags.ProcesosPara) = 0 Then Exit Sub
        
        ' Prevengo avivadas
        data = Replace$(data, "*:*", vbNullString)
        
        ' Anteponemos el nombre del user
        data = .name & "*:*" & data

        Dim ListaGMs() As String
        ListaGMs = Split(.flags.ProcesosPara, ":")
        
        Dim i As Integer, tGM As Integer

        For i = LBound(ListaGMs) To UBound(ListaGMs)
            tGM = NameIndex(ListaGMs(i))
            
            If tGM > 0 Then
                Call WriteShowProcesses(tGM, data)

            End If

        Next
        
        .flags.ProcesosPara = vbNullString

    End With
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleProcesses", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleGetMapInfo(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If EsGM(UserIndex) Then
            
            Dim Response As String
            
            Response = "[Info de mapa " & .Pos.Map & "]" & vbNewLine
            Response = Response & "Nombre = " & MapInfo(.Pos.Map).map_name & vbNewLine
            Response = Response & "Seguro = " & MapInfo(.Pos.Map).Seguro & vbNewLine
            Response = Response & "Newbie = " & MapInfo(.Pos.Map).Newbie & vbNewLine
            Response = Response & "Nivel = " & MapInfo(.Pos.Map).MinLevel & "/" & MapInfo(.Pos.Map).MaxLevel & vbNewLine
            Response = Response & "SinInviOcul = " & MapInfo(.Pos.Map).SinInviOcul & vbNewLine
            Response = Response & "SinMagia = " & MapInfo(.Pos.Map).SinMagia & vbNewLine
            Response = Response & "SoloClanes = " & MapInfo(.Pos.Map).SoloClanes & vbNewLine
            Response = Response & "NoPKs = " & MapInfo(.Pos.Map).NoPKs & vbNewLine
            Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
            Response = Response & "Salida = " & MapInfo(.Pos.Map).Salida.Map & "-" & MapInfo(.Pos.Map).Salida.X & "-" & MapInfo(.Pos.Map).Salida.Y & vbNewLine
            Response = Response & "Terreno = " & MapInfo(.Pos.Map).terrain & vbNewLine
            Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
            Response = Response & "Zona = " & MapInfo(.Pos.Map).zone & vbNewLine
            
            Call WriteConsoleMsg(UserIndex, Response, FontTypeNames.FONTTYPE_INFO)
        
        End If
    
    End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim name As String
        name = .incomingData.ReadASCIIString()

        If LenB(name) = 0 Then Exit Sub

        If EsGmChar(name) Then
            Call WriteConsoleMsg(UserIndex, "No podés denunciar a un administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim tUser As Integer
        tUser = NameIndex(name)
        
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim Denuncia As String, HayChat As Boolean
        Denuncia = "[Últimos mensajes de " & UserList(tUser).name & "]" & vbNewLine
        
        Dim i As Integer

        For i = 1 To UBound(UserList(tUser).flags.ChatHistory)

            If LenB(UserList(tUser).flags.ChatHistory(i)) <> 0 Then
                Denuncia = Denuncia & UserList(tUser).flags.ChatHistory(i) & vbNewLine
                HayChat = True

            End If

        Next
        
        If Not HayChat Then
            Call WriteConsoleMsg(UserIndex, "El usuario no ha escrito nada. Recordá que las denuncias inválidas pueden ser motivo de advertencia.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(UserIndex).donador.activo = 1 Then
            Call Ayuda.Push(.name, Denuncia, "Denuncia a " & UserList(tUser).name & "-Prioritario")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido una nueva denuncia de parte de " & .name & "(Prioritario).", FontTypeNames.FONTTYPE_SERVER))
        
        Else
            Call Ayuda.Push(.name, Denuncia, "Denuncia a " & UserList(tUser).name)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido una nueva denuncia de parte de " & .name & ".", FontTypeNames.FONTTYPE_SERVER))

        End If

        Call WriteConsoleMsg(UserIndex, "Tu denuncia fue recibida por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)

        Call LogConsulta(.name & " (Denuncia a " & UserList(tUser).name & ")" & vbNewLine & Denuncia)

    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Private Sub HandleSeguroResu(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        .flags.SeguroResu = Not .flags.SeguroResu
        
        Call WriteSeguroResu(UserIndex, .flags.SeguroResu)
    
    End With

End Sub

Private Sub HandleCuentaExtractItem(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCuentaExtractItem_Err

    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Retirar item de cuenta
    '***************************************************

    With UserList(UserIndex)

        Dim slot        As Byte

        Dim slotdestino As Byte

        Dim amount      As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        slotdestino = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.TargetNPC < 1 Then Exit Sub
        
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
        
        'acá va el guardado en memoria
        
        'User retira el item del slot
        'Call UserRetiraItem(UserIndex, slot, Amount, slotdestino)

    End With
        
    Exit Sub

HandleCuentaExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCuentaExtractItem", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleCuentaDeposit(ByVal UserIndex As Integer)
        
    On Error GoTo HandleCuentaDeposit_Err

    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Depositar item en cuenta
    '***************************************************
    
    With UserList(UserIndex)

        Dim slot        As Byte

        Dim slotdestino As Byte

        Dim amount      As Integer
        
        slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        slotdestino = .incomingData.ReadByte()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'íEl target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
            
        If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'acá va el guardado en memoria
            
        'User deposita el item del slot rdata
        'Call UserDepositaItem(UserIndex, slot, Amount, slotdestino)

    End With
        
    Exit Sub

HandleCuentaDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCuentaDeposit", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Private Sub HandleCommerceSendChatMessage(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim chatMessage As String
        
        chatMessage = "[" & UserList(UserIndex).name & "] " & .incomingData.ReadASCIIString
        
        'El mensaje se lo envío al destino
        Call WriteCommerceRecieveChatMessage(UserList(UserIndex).ComUsu.DestUsu, chatMessage)
        
        'y tambien a mi mismo
        Call WriteCommerceRecieveChatMessage(UserIndex, chatMessage)

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceSendChatMessage", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub HandleLogMacroClickHechizo(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("AntiCheat> El usuario " & .name & " se le cerró el cliente por posible uso de macro de hechizos", FontTypeNames.FONTTYPE_INFO))
        Call LogHackAttemp("Usuario: " & .name & "   " & "Ip: " & .ip & " Posible uso de macro de hechizos.")

    End With

End Sub

Private Sub HandleCreateEvent(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim name As String
        name = .incomingData.ReadASCIIString()

        If LenB(name) = 0 Then Exit Sub
    
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
    
        Select Case UCase$(name)

            Case "INVASION BANDER"
                Call IniciarEvento(TipoEvento.Invasion, 1)
                Call LogGM(.name, "Forzó el evento Invasión en Banderbille.")
                
            Case "INVASION CARCEL"
                Call IniciarEvento(TipoEvento.Invasion, 2)
                Call LogGM(.name, "Forzó el evento Invasión en Carcel.")

            Case Else
                Call WriteConsoleMsg(UserIndex, "No existe el evento """ & name & """.", FontTypeNames.FONTTYPE_INFO)

        End Select

    End With
    
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateEvent", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

