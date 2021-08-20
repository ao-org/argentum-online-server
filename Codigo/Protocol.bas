Attribute VB_Name = "Protocol"
Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR             As String * 1 = vbNullChar

Public Enum ServerPacketID

    Connected
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
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
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
    'SendNight              ' NOC
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
    BindKeys
    ShowFrmLogear
    ShowFrmMapa
    InmovilizadoOK
    BarFx
    SetEscribiendo
    LocaleMsg
    ShowPregunta
    DatosGrupo
    ubicacion
    ArmaMov
    EscudoMov
    ViajarForm
    Oxigeno
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    CharUpdateMAN
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
    SeguroResu
    Stopped
    InvasionInfo
    CommerceRecieveChatMessage
    DoAnimation
    OpenCrafting
    CraftingItem
    CraftingCatalyst
    CraftingResult
    ForceUpdate
    GuardNotice
    ObjQuestListSend
    [PacketCount]
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
    Change_Heading           'CHEA
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
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    punishments             '/PENAS
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
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    banip                   '/BANIP
    UnBanIp                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    Tile_BlockedToggle       '/BLOQ
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
    Participar              '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
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
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    
    'Nuevas Ladder
    SetSpeed
    GlobalMessage           '/CONSOLA
    GlobalOnOff
    IngresarConCuenta
    BorrarPJ
    DarLlaveAUsuario
    SacarLlave
    VerLlaves
    UseKey
    Day
    SetTime
    DonateGold              '/DONAR
    Promedio                '/PROMEDIO
    GiveItem                '/DAR
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
    InvitarGrupo
    ResponderPregunta
    RequestGrupo
    AbandonarGrupo
    HecharDeGrupo
    MacroPossent
    SubastaInfo
    BanCuenta
    UnbanCuenta
    CerrarCliente
    EventoInfo
    CrearEvento
    BanTemporal
    CancelarExit
    CrearTorneo
    ComenzarTorneo
    CancelarTorneo
    BusquedaTesoro
    CompletarViaje
    BovedaMoveItem
    QuieroFundarClan
    llamadadeclan
    MarcaDeClanPack
    MarcaDeGMPack
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    SeguroClan
    CreatePretorianClan     '/CREARPRETORIANOS
    Home                    '/HOGAR
    Consulta                '/CONSULTA
    GetMapInfo              '/MAPINFO
    FinEvento
    SeguroResu
    CuentaExtractItem
    CuentaDeposit
    CreateEvent
    CommerceSendChatMessage
    LogMacroClickHechizo
    AddItemCrafting
    RemoveItemCrafting
    AddCatalyst
    RemoveCatalyst
    CraftItem
    CloseCrafting
    MoveCraftItem
    PetLeaveAll
    GuardNoticeResponse
    GuardResendVerificationCode
    ResetChar               '/RESET NICK
    DeleteItem
    
    [PacketCount]
End Enum

Public Enum e_EditOptions

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
    eo_CASCO
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

Public Enum e_FontTypeNames

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

Private Reader  As Network.Reader

Public Sub InitializePacketList()
    Call Protocol_Writes.InitializeAuxiliaryBuffer
End Sub

''
' Handles incoming data.
'
' @param    UserIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer, ByVal Message As Network.Reader) As Boolean

On Error Resume Next

    Set Reader = Message
    
    Dim PacketID As Long:
    PacketID = Reader.ReadInt

    'Does the packet requires a logged user??
    If Not (PacketID = ClientPacketID.LoginExistingChar Or _
            PacketID = ClientPacketID.LoginNewChar Or _
            PacketID = ClientPacketID.IngresarConCuenta Or _
            PacketID = ClientPacketID.BorrarPJ Or _
            PacketID = ClientPacketID.ThrowDice Or _
            PacketID = ClientPacketID.GuardNoticeResponse) Then
               
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf PacketID <= ClientPacketID.[PacketCount] Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf PacketID <= ClientPacketID.[PacketCount] Then
        UserList(UserIndex).Counters.IdleCount = 0
    End If
    
    Select Case PacketID
        Case ClientPacketID.LoginExistingChar
            Call HandleLoginExistingChar(UserIndex)
        Case ClientPacketID.LoginNewChar
            Call HandleLoginNewChar(UserIndex)
        Case ClientPacketID.ThrowDice
            Call HandleThrowDice(UserIndex)
        Case ClientPacketID.Talk
            Call HandleTalk(UserIndex)
        Case ClientPacketID.Yell
            Call HandleYell(UserIndex)
        Case ClientPacketID.Whisper
            Call HandleWhisper(UserIndex)
        Case ClientPacketID.Walk
            Call HandleWalk(UserIndex)
        Case ClientPacketID.RequestPositionUpdate
            Call HandleRequestPositionUpdate(UserIndex)
        Case ClientPacketID.Attack
            Call HandleAttack(UserIndex)
        Case ClientPacketID.PickUp
            Call HandlePickUp(UserIndex)
        Case ClientPacketID.SafeToggle
            Call HandleSafeToggle(UserIndex)
        Case ClientPacketID.PartySafeToggle
            Call HandlePartyToggle(UserIndex)
        Case ClientPacketID.RequestGuildLeaderInfo
            Call HandleRequestGuildLeaderInfo(UserIndex)
        Case ClientPacketID.RequestAtributes
            Call HandleRequestAtributes(UserIndex)
        Case ClientPacketID.RequestSkills
            Call HandleRequestSkills(UserIndex)
        Case ClientPacketID.RequestMiniStats
            Call HandleRequestMiniStats(UserIndex)
        Case ClientPacketID.CommerceEnd
            Call HandleCommerceEnd(UserIndex)
        Case ClientPacketID.UserCommerceEnd
            Call HandleUserCommerceEnd(UserIndex)
        Case ClientPacketID.BankEnd
            Call HandleBankEnd(UserIndex)
        Case ClientPacketID.UserCommerceOk
            Call HandleUserCommerceOk(UserIndex)
        Case ClientPacketID.UserCommerceReject
            Call HandleUserCommerceReject(UserIndex)
        Case ClientPacketID.Drop
            Call HandleDrop(UserIndex)
        Case ClientPacketID.CastSpell
            Call HandleCastSpell(UserIndex)
        Case ClientPacketID.LeftClick
            Call HandleLeftClick(UserIndex)
        Case ClientPacketID.DoubleClick
            Call HandleDoubleClick(UserIndex)
        Case ClientPacketID.Work
            Call HandleWork(UserIndex)
        Case ClientPacketID.UseSpellMacro
            Call HandleUseSpellMacro(UserIndex)
        Case ClientPacketID.UseItem
            Call HandleUseItem(UserIndex)
        Case ClientPacketID.CraftBlacksmith
            Call HandleCraftBlacksmith(UserIndex)
        Case ClientPacketID.CraftCarpenter
            Call HandleCraftCarpenter(UserIndex)
        Case ClientPacketID.WorkLeftClick
            Call HandleWorkLeftClick(UserIndex)
        Case ClientPacketID.CreateNewGuild
            Call HandleCreateNewGuild(UserIndex)
        Case ClientPacketID.SpellInfo
            Call HandleSpellInfo(UserIndex)
        Case ClientPacketID.EquipItem
            Call HandleEquipItem(UserIndex)
        Case ClientPacketID.Change_Heading
            Call HandleChange_Heading(UserIndex)
        Case ClientPacketID.ModifySkills
            Call HandleModifySkills(UserIndex)
        Case ClientPacketID.Train
            Call HandleTrain(UserIndex)
        Case ClientPacketID.CommerceBuy
            Call HandleCommerceBuy(UserIndex)
        Case ClientPacketID.BankExtractItem
            Call HandleBankExtractItem(UserIndex)
        Case ClientPacketID.CommerceSell
            Call HandleCommerceSell(UserIndex)
        Case ClientPacketID.BankDeposit
            Call HandleBankDeposit(UserIndex)
        Case ClientPacketID.ForumPost
            Call HandleForumPost(UserIndex)
        Case ClientPacketID.MoveSpell
            Call HandleMoveSpell(UserIndex)
        Case ClientPacketID.ClanCodexUpdate
            Call HandleClanCodexUpdate(UserIndex)
        Case ClientPacketID.UserCommerceOffer
            Call HandleUserCommerceOffer(UserIndex)
        Case ClientPacketID.GuildAcceptPeace
            Call HandleGuildAcceptPeace(UserIndex)
        Case ClientPacketID.GuildRejectAlliance
            Call HandleGuildRejectAlliance(UserIndex)
        Case ClientPacketID.GuildRejectPeace
            Call HandleGuildRejectPeace(UserIndex)
        Case ClientPacketID.GuildAcceptAlliance
            Call HandleGuildAcceptAlliance(UserIndex)
        Case ClientPacketID.GuildOfferPeace
            Call HandleGuildOfferPeace(UserIndex)
        Case ClientPacketID.GuildOfferAlliance
            Call HandleGuildOfferAlliance(UserIndex)
        Case ClientPacketID.GuildAllianceDetails
            Call HandleGuildAllianceDetails(UserIndex)
        Case ClientPacketID.GuildPeaceDetails
            Call HandleGuildPeaceDetails(UserIndex)
        Case ClientPacketID.GuildRequestJoinerInfo
            Call HandleGuildRequestJoinerInfo(UserIndex)
        Case ClientPacketID.GuildAlliancePropList
            Call HandleGuildAlliancePropList(UserIndex)
        Case ClientPacketID.GuildPeacePropList
            Call HandleGuildPeacePropList(UserIndex)
        Case ClientPacketID.GuildDeclareWar
            Call HandleGuildDeclareWar(UserIndex)
        Case ClientPacketID.GuildNewWebsite
            Call HandleGuildNewWebsite(UserIndex)
        Case ClientPacketID.GuildAcceptNewMember
            Call HandleGuildAcceptNewMember(UserIndex)
        Case ClientPacketID.GuildRejectNewMember
            Call HandleGuildRejectNewMember(UserIndex)
        Case ClientPacketID.GuildKickMember
            Call HandleGuildKickMember(UserIndex)
        Case ClientPacketID.GuildUpdateNews
            Call HandleGuildUpdateNews(UserIndex)
        Case ClientPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo(UserIndex)
        Case ClientPacketID.GuildOpenElections
            Call HandleGuildOpenElections(UserIndex)
        Case ClientPacketID.GuildRequestMembership
            Call HandleGuildRequestMembership(UserIndex)
        Case ClientPacketID.GuildRequestDetails
            Call HandleGuildRequestDetails(UserIndex)
        Case ClientPacketID.Online
            Call HandleOnline(UserIndex)
        Case ClientPacketID.Quit
            Call HandleQuit(UserIndex)
        Case ClientPacketID.GuildLeave
            Call HandleGuildLeave(UserIndex)
        Case ClientPacketID.RequestAccountState
            Call HandleRequestAccountState(UserIndex)
        Case ClientPacketID.PetStand
            Call HandlePetStand(UserIndex)
        Case ClientPacketID.PetFollow
            Call HandlePetFollow(UserIndex)
        Case ClientPacketID.PetLeave
            Call HandlePetLeave(UserIndex)
        Case ClientPacketID.GrupoMsg
            Call HandleGrupoMsg(UserIndex)
        Case ClientPacketID.TrainList
            Call HandleTrainList(UserIndex)
        Case ClientPacketID.Rest
            Call HandleRest(UserIndex)
        Case ClientPacketID.Meditate
            Call HandleMeditate(UserIndex)
        Case ClientPacketID.Resucitate
            Call HandleResucitate(UserIndex)
        Case ClientPacketID.Heal
            Call HandleHeal(UserIndex)
        Case ClientPacketID.Help
            Call HandleHelp(UserIndex)
        Case ClientPacketID.RequestStats
            Call HandleRequestStats(UserIndex)
        Case ClientPacketID.CommerceStart
            Call HandleCommerceStart(UserIndex)
        Case ClientPacketID.BankStart
            Call HandleBankStart(UserIndex)
        Case ClientPacketID.Enlist
            Call HandleEnlist(UserIndex)
        Case ClientPacketID.Information
            Call HandleInformation(UserIndex)
        Case ClientPacketID.Reward
            Call HandleReward(UserIndex)
        Case ClientPacketID.RequestMOTD
            Call HandleRequestMOTD(UserIndex)
        Case ClientPacketID.UpTime
            Call HandleUpTime(UserIndex)
        Case ClientPacketID.Inquiry
            Call HandleInquiry(UserIndex)
        Case ClientPacketID.GuildMessage
            Call HandleGuildMessage(UserIndex)
        Case ClientPacketID.CentinelReport
            Call HandleCentinelReport(UserIndex)
        Case ClientPacketID.GuildOnline
            Call HandleGuildOnline(UserIndex)
        Case ClientPacketID.CouncilMessage
            Call HandleCouncilMessage(UserIndex)
        Case ClientPacketID.RoleMasterRequest
            Call HandleRoleMasterRequest(UserIndex)
        Case ClientPacketID.ChangeDescription
            Call HandleChangeDescription(UserIndex)
        Case ClientPacketID.GuildVote
            Call HandleGuildVote(UserIndex)
        Case ClientPacketID.punishments
            Call HandlePunishments(UserIndex)
        Case ClientPacketID.Gamble
            Call HandleGamble(UserIndex)
        Case ClientPacketID.InquiryVote
            Call HandleInquiryVote(UserIndex)
        Case ClientPacketID.LeaveFaction
            Call HandleLeaveFaction(UserIndex)
        Case ClientPacketID.BankExtractGold
            Call HandleBankExtractGold(UserIndex)
        Case ClientPacketID.BankDepositGold
            Call HandleBankDepositGold(UserIndex)
        Case ClientPacketID.Denounce
            Call HandleDenounce(UserIndex)
        Case ClientPacketID.Ping
            Call HandlePing(UserIndex)
        Case ClientPacketID.GMMessage
            Call HandleGMMessage(UserIndex)
        Case ClientPacketID.showName
            Call HandleShowName(UserIndex)
        Case ClientPacketID.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        Case ClientPacketID.OnlineChaosLegion
            Call HandleOnlineChaosLegion(UserIndex)
        Case ClientPacketID.GoNearby
            Call HandleGoNearby(UserIndex)
        Case ClientPacketID.comment
            Call HandleComment(UserIndex)
        Case ClientPacketID.serverTime
            Call HandleServerTime(UserIndex)
        Case ClientPacketID.Where
            Call HandleWhere(UserIndex)
        Case ClientPacketID.CreaturesInMap
            Call HandleCreaturesInMap(UserIndex)
        Case ClientPacketID.WarpMeToTarget
            Call HandleWarpMeToTarget(UserIndex)
        Case ClientPacketID.WarpChar
            Call HandleWarpChar(UserIndex)
        Case ClientPacketID.Silence
            Call HandleSilence(UserIndex)
        Case ClientPacketID.SOSShowList
            Call HandleSOSShowList(UserIndex)
        Case ClientPacketID.SOSRemove
            Call HandleSOSRemove(UserIndex)
        Case ClientPacketID.GoToChar
            Call HandleGoToChar(UserIndex)
        Case ClientPacketID.invisible
            Call HandleInvisible(UserIndex)
        Case ClientPacketID.GMPanel
            Call HandleGMPanel(UserIndex)
        Case ClientPacketID.RequestUserList
            Call HandleRequestUserList(UserIndex)
        Case ClientPacketID.Working
            Call HandleWorking(UserIndex)
        Case ClientPacketID.Hiding
            Call HandleHiding(UserIndex)
        Case ClientPacketID.Jail
            Call HandleJail(UserIndex)
        Case ClientPacketID.KillNPC
            Call HandleKillNPC(UserIndex)
        Case ClientPacketID.WarnUser
            Call HandleWarnUser(UserIndex)
        Case ClientPacketID.EditChar
            Call HandleEditChar(UserIndex)
        Case ClientPacketID.RequestCharInfo
            Call HandleRequestCharInfo(UserIndex)
        Case ClientPacketID.RequestCharStats
            Call HandleRequestCharStats(UserIndex)
        Case ClientPacketID.RequestCharGold
            Call HandleRequestCharGold(UserIndex)
        Case ClientPacketID.RequestCharInventory
            Call HandleRequestCharInventory(UserIndex)
        Case ClientPacketID.RequestCharBank
            Call HandleRequestCharBank(UserIndex)
        Case ClientPacketID.RequestCharSkills
            Call HandleRequestCharSkills(UserIndex)
        Case ClientPacketID.ReviveChar
            Call HandleReviveChar(UserIndex)
        Case ClientPacketID.OnlineGM
            Call HandleOnlineGM(UserIndex)
        Case ClientPacketID.OnlineMap
            Call HandleOnlineMap(UserIndex)
        Case ClientPacketID.Forgive
            Call HandleForgive(UserIndex)
        Case ClientPacketID.Kick
            Call HandleKick(UserIndex)
        Case ClientPacketID.Execute
            Call HandleExecute(UserIndex)
        Case ClientPacketID.BanChar
            Call HandleBanChar(UserIndex)
        Case ClientPacketID.UnbanChar
            Call HandleUnbanChar(UserIndex)
        Case ClientPacketID.NPCFollow
            Call HandleNPCFollow(UserIndex)
        Case ClientPacketID.SummonChar
            Call HandleSummonChar(UserIndex)
        Case ClientPacketID.SpawnListRequest
            Call HandleSpawnListRequest(UserIndex)
        Case ClientPacketID.SpawnCreature
            Call HandleSpawnCreature(UserIndex)
        Case ClientPacketID.ResetNPCInventory
            Call HandleResetNPCInventory(UserIndex)
        Case ClientPacketID.CleanWorld
            Call HandleCleanWorld(UserIndex)
        Case ClientPacketID.ServerMessage
            Call HandleServerMessage(UserIndex)
        Case ClientPacketID.NickToIP
            Call HandleNickToIP(UserIndex)
        Case ClientPacketID.IPToNick
            Call HandleIPToNick(UserIndex)
        Case ClientPacketID.GuildOnlineMembers
            Call HandleGuildOnlineMembers(UserIndex)
        Case ClientPacketID.TeleportCreate
            Call HandleTeleportCreate(UserIndex)
        Case ClientPacketID.TeleportDestroy
            Call HandleTeleportDestroy(UserIndex)
        Case ClientPacketID.RainToggle
            Call HandleRainToggle(UserIndex)
        Case ClientPacketID.SetCharDescription
            Call HandleSetCharDescription(UserIndex)
        Case ClientPacketID.ForceMIDIToMap
            Call HandleForceMIDIAll(UserIndex)
        Case ClientPacketID.ForceWAVEToMap
            Call HandleForceWAVEToMap(UserIndex)
        Case ClientPacketID.RoyalArmyMessage
            Call HandleRoyalArmyMessage(UserIndex)
        Case ClientPacketID.ChaosLegionMessage
            Call HandleChaosLegionMessage(UserIndex)
        Case ClientPacketID.CitizenMessage
            Call HandleCitizenMessage(UserIndex)
        Case ClientPacketID.CriminalMessage
            Call HandleCriminalMessage(UserIndex)
        Case ClientPacketID.TalkAsNPC
            Call HandleTalkAsNPC(UserIndex)
        Case ClientPacketID.DestroyAllItemsInArea
            Call HandleDestroyAllItemsInArea(UserIndex)
        Case ClientPacketID.AcceptRoyalCouncilMember
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        Case ClientPacketID.AcceptChaosCouncilMember
            Call HandleAcceptChaosCouncilMember(UserIndex)
        Case ClientPacketID.ItemsInTheFloor
            Call HandleItemsInTheFloor(UserIndex)
        Case ClientPacketID.MakeDumb
            Call HandleMakeDumb(UserIndex)
        Case ClientPacketID.MakeDumbNoMore
            Call HandleMakeDumbNoMore(UserIndex)
        Case ClientPacketID.CouncilKick
            Call HandleCouncilKick(UserIndex)
        Case ClientPacketID.SetTrigger
            Call HandleSetTrigger(UserIndex)
        Case ClientPacketID.AskTrigger
            Call HandleAskTrigger(UserIndex)
        Case ClientPacketID.BannedIPList
            Call HandleBannedIPList(UserIndex)
        Case ClientPacketID.BannedIPReload
            Call HandleBannedIPReload(UserIndex)
        Case ClientPacketID.GuildMemberList
            Call HandleGuildMemberList(UserIndex)
        Case ClientPacketID.GuildBan
            Call HandleGuildBan(UserIndex)
        Case ClientPacketID.banip
            Call HandleBanIP(UserIndex)
        Case ClientPacketID.UnBanIp
            Call HandleUnbanIP(UserIndex)
        Case ClientPacketID.CreateItem
            Call HandleCreateItem(UserIndex)
        Case ClientPacketID.DestroyItems
            Call HandleDestroyItems(UserIndex)
        Case ClientPacketID.ChaosLegionKick
            Call HandleChaosLegionKick(UserIndex)
        Case ClientPacketID.RoyalArmyKick
            Call HandleRoyalArmyKick(UserIndex)
        Case ClientPacketID.ForceMIDIAll
            Call HandleForceMIDIAll(UserIndex)
        Case ClientPacketID.ForceWAVEAll
            Call HandleForceWAVEAll(UserIndex)
        Case ClientPacketID.RemovePunishment
            Call HandleRemovePunishment(UserIndex)
        Case ClientPacketID.Tile_BlockedToggle
            Call HandleTile_BlockedToggle(UserIndex)
        Case ClientPacketID.KillNPCNoRespawn
            Call HandleKillNPCNoRespawn(UserIndex)
        Case ClientPacketID.KillAllNearbyNPCs
            Call HandleKillAllNearbyNPCs(UserIndex)
        Case ClientPacketID.LastIP
            Call HandleLastIP(UserIndex)
        Case ClientPacketID.ChangeMOTD
            Call HandleChangeMOTD(UserIndex)
        Case ClientPacketID.SetMOTD
            Call HandleSetMOTD(UserIndex)
        Case ClientPacketID.SystemMessage
            Call HandleSystemMessage(UserIndex)
        Case ClientPacketID.CreateNPC
            Call HandleCreateNPC(UserIndex)
        Case ClientPacketID.CreateNPCWithRespawn
            Call HandleCreateNPCWithRespawn(UserIndex)
        Case ClientPacketID.ImperialArmour
            Call HandleImperialArmour(UserIndex)
        Case ClientPacketID.ChaosArmour
            Call HandleChaosArmour(UserIndex)
        Case ClientPacketID.NavigateToggle
            Call HandleNavigateToggle(UserIndex)
        Case ClientPacketID.ServerOpenToUsersToggle
            Call HandleServerOpenToUsersToggle(UserIndex)
        Case ClientPacketID.Participar
            Call HandleParticipar(UserIndex)
        Case ClientPacketID.TurnCriminal
            Call HandleTurnCriminal(UserIndex)
        Case ClientPacketID.ResetFactions
            Call HandleResetFactions(UserIndex)
        Case ClientPacketID.RemoveCharFromGuild
            Call HandleRemoveCharFromGuild(UserIndex)
        Case ClientPacketID.AlterName
            Call HandleAlterName(UserIndex)
        Case ClientPacketID.DoBackUp
            Call HandleDoBackUp(UserIndex)
        Case ClientPacketID.ShowGuildMessages
            Call HandleShowGuildMessages(UserIndex)
        Case ClientPacketID.SaveMap
            Call HandleSaveMap(UserIndex)
        Case ClientPacketID.ChangeMapInfoPK
            Call HandleChangeMapInfoPK(UserIndex)
        Case ClientPacketID.ChangeMapInfoBackup
            Call HandleChangeMapInfoBackup(UserIndex)
        Case ClientPacketID.ChangeMapInfoRestricted
            Call HandleChangeMapInfoRestricted(UserIndex)
        Case ClientPacketID.ChangeMapInfoNoMagic
            Call HandleChangeMapInfoNoMagic(UserIndex)
        Case ClientPacketID.ChangeMapInfoNoInvi
            Call HandleChangeMapInfoNoInvi(UserIndex)
        Case ClientPacketID.ChangeMapInfoNoResu
            Call HandleChangeMapInfoNoResu(UserIndex)
        Case ClientPacketID.ChangeMapInfoLand
            Call HandleChangeMapInfoLand(UserIndex)
        Case ClientPacketID.ChangeMapInfoZone
            Call HandleChangeMapInfoZone(UserIndex)
        Case ClientPacketID.SaveChars
            Call HandleSaveChars(UserIndex)
        Case ClientPacketID.CleanSOS
            Call HandleCleanSOS(UserIndex)
        Case ClientPacketID.ShowServerForm
            Call HandleShowServerForm(UserIndex)
        Case ClientPacketID.night
            Call HandleNight(UserIndex)
        Case ClientPacketID.KickAllChars
            Call HandleKickAllChars(UserIndex)
        Case ClientPacketID.ReloadNPCs
            Call HandleReloadNPCs(UserIndex)
        Case ClientPacketID.ReloadServerIni
            Call HandleReloadServerIni(UserIndex)
        Case ClientPacketID.ReloadSpells
            Call HandleReloadSpells(UserIndex)
        Case ClientPacketID.ReloadObjects
            Call HandleReloadObjects(UserIndex)
        Case ClientPacketID.Restart
            Call HandleRestart(UserIndex)
        Case ClientPacketID.ChatColor
            Call HandleChatColor(UserIndex)
        Case ClientPacketID.Ignored
            Call HandleIgnored(UserIndex)
        Case ClientPacketID.CheckSlot
            Call HandleCheckSlot(UserIndex)
        Case ClientPacketID.SetSpeed
            Call HandleSetSpeed(UserIndex)
        Case ClientPacketID.GlobalMessage
            Call HandleGlobalMessage(UserIndex)
        Case ClientPacketID.GlobalOnOff
            Call HandleGlobalOnOff(UserIndex)
        Case ClientPacketID.IngresarConCuenta
            Call HandleIngresarConCuenta(UserIndex)
        Case ClientPacketID.BorrarPJ
            Call HandleBorrarPJ(UserIndex)
        Case ClientPacketID.DarLlaveAUsuario
            Call HandleDarLlaveAUsuario(UserIndex)
        Case ClientPacketID.SacarLlave
            Call HandleSacarLlave(UserIndex)
        Case ClientPacketID.VerLlaves
            Call HandleVerLlaves(UserIndex)
        Case ClientPacketID.UseKey
            Call HandleUseKey(UserIndex)
        Case ClientPacketID.Day
            Call HandleDay(UserIndex)
        Case ClientPacketID.SetTime
            Call HandleSetTime(UserIndex)
        Case ClientPacketID.DonateGold
            Call HandleDonateGold(UserIndex)
        Case ClientPacketID.Promedio
            Call HandlePromedio(UserIndex)
        Case ClientPacketID.GiveItem
            Call HandleGiveItem(UserIndex)
        Case ClientPacketID.OfertaInicial
            Call HandleOfertaInicial(UserIndex)
        Case ClientPacketID.OfertaDeSubasta
            Call HandleOfertaDeSubasta(UserIndex)
        Case ClientPacketID.QuestionGM
            Call HandleQuestionGM(UserIndex)
        Case ClientPacketID.CuentaRegresiva
            Call HandleCuentaRegresiva(UserIndex)
        Case ClientPacketID.PossUser
            Call HandlePossUser(UserIndex)
        Case ClientPacketID.Duel
            Call HandleDuel(UserIndex)
        Case ClientPacketID.AcceptDuel
            Call HandleAcceptDuel(UserIndex)
        Case ClientPacketID.CancelDuel
            Call HandleCancelDuel(UserIndex)
        Case ClientPacketID.QuitDuel
            Call HandleQuitDuel(UserIndex)
        Case ClientPacketID.NieveToggle
            Call HandleNieveToggle(UserIndex)
        Case ClientPacketID.NieblaToggle
            Call HandleNieblaToggle(UserIndex)
        Case ClientPacketID.TransFerGold
            Call HandleTransFerGold(UserIndex)
        Case ClientPacketID.Moveitem
            Call HandleMoveItem(UserIndex)
        Case ClientPacketID.Genio
            Call HandleGenio(UserIndex)
        Case ClientPacketID.Casarse
            Call HandleCasamiento(UserIndex)
        Case ClientPacketID.CraftAlquimista
            Call HandleCraftAlquimia(UserIndex)
        Case ClientPacketID.RequestFamiliar
            Call HandleRequestFamiliar(UserIndex)
        Case ClientPacketID.FlagTrabajar
            Call HandleFlagTrabajar(UserIndex)
        Case ClientPacketID.CraftSastre
            Call HandleCraftSastre(UserIndex)
        Case ClientPacketID.MensajeUser
            Call HandleMensajeUser(UserIndex)
        Case ClientPacketID.TraerBoveda
            Call HandleTraerBoveda(UserIndex)
        Case ClientPacketID.CompletarAccion
            Call HandleCompletarAccion(UserIndex)
        Case ClientPacketID.Escribiendo
            Call HandleEscribiendo(UserIndex)
        Case ClientPacketID.InvitarGrupo
            Call HandleInvitarGrupo(UserIndex)
        Case ClientPacketID.ResponderPregunta
            Call HandleResponderPregunta(UserIndex)
        Case ClientPacketID.RequestGrupo
            Call HandleRequestGrupo(UserIndex)
        Case ClientPacketID.AbandonarGrupo
            Call HandleAbandonarGrupo(UserIndex)
        Case ClientPacketID.HecharDeGrupo
            Call HandleHecharDeGrupo(UserIndex)
        Case ClientPacketID.MacroPossent
            Call HandleMacroPos(UserIndex)
        Case ClientPacketID.SubastaInfo
            Call HandleSubastaInfo(UserIndex)
        Case ClientPacketID.BanCuenta
            Call HandleBanCuenta(UserIndex)
        Case ClientPacketID.UnbanCuenta
            Call HandleUnBanCuenta(UserIndex)
        Case ClientPacketID.CerrarCliente
            Call HandleCerrarCliente(UserIndex)
        Case ClientPacketID.EventoInfo
            Call HandleEventoInfo(UserIndex)
        Case ClientPacketID.CrearEvento
            Call HandleCrearEvento(UserIndex)
        Case ClientPacketID.BanTemporal
            Call HandleBanTemporal(UserIndex)
        Case ClientPacketID.CancelarExit
            Call HandleCancelarExit(UserIndex)
        Case ClientPacketID.CrearTorneo
            Call HandleCrearTorneo(UserIndex)
        Case ClientPacketID.ComenzarTorneo
            Call HandleComenzarTorneo(UserIndex)
        Case ClientPacketID.CancelarTorneo
            Call HandleCancelarTorneo(UserIndex)
        Case ClientPacketID.BusquedaTesoro
            Call HandleBusquedaTesoro(UserIndex)
        Case ClientPacketID.CompletarViaje
            Call HandleCompletarViaje(UserIndex)
        Case ClientPacketID.BovedaMoveItem
            Call HandleBovedaMoveItem(UserIndex)
        Case ClientPacketID.QuieroFundarClan
            Call HandleQuieroFundarClan(UserIndex)
        Case ClientPacketID.llamadadeclan
            Call HandleLlamadadeClan(UserIndex)
        Case ClientPacketID.MarcaDeClanPack
            Call HandleMarcaDeClan(UserIndex)
        Case ClientPacketID.MarcaDeGMPack
            Call HandleMarcaDeGM(UserIndex)
        Case ClientPacketID.Quest
            Call HandleQuest(UserIndex)
        Case ClientPacketID.QuestAccept
            Call HandleQuestAccept(UserIndex)
        Case ClientPacketID.QuestListRequest
            Call HandleQuestListRequest(UserIndex)
        Case ClientPacketID.QuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)
        Case ClientPacketID.QuestAbandon
            Call HandleQuestAbandon(UserIndex)
        Case ClientPacketID.SeguroClan
            Call HandleSeguroClan(UserIndex)
        Case ClientPacketID.CreatePretorianClan
            Call HandleCreatePretorianClan(UserIndex)
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
        Case ClientPacketID.Consulta
            Call HandleConsulta(UserIndex)
        Case ClientPacketID.GetMapInfo
            Call HandleGetMapInfo(UserIndex)
        Case ClientPacketID.FinEvento
            Call HandleFinEvento(UserIndex)
        Case ClientPacketID.SeguroResu
            Call HandleSeguroResu(UserIndex)
        Case ClientPacketID.CuentaExtractItem
            Call HandleCuentaExtractItem(UserIndex)
        Case ClientPacketID.CuentaDeposit
            Call HandleCuentaDeposit(UserIndex)
        Case ClientPacketID.CreateEvent
            Call HandleCreateEvent(UserIndex)
        Case ClientPacketID.CommerceSendChatMessage
            Call HandleCommerceSendChatMessage(UserIndex)
        Case ClientPacketID.LogMacroClickHechizo
            Call HandleLogMacroClickHechizo(UserIndex)
        Case ClientPacketID.AddItemCrafting
            Call HandleAddItemCrafting(UserIndex)
        Case ClientPacketID.RemoveItemCrafting
            Call HandleRemoveItemCrafting(UserIndex)
        Case ClientPacketID.AddCatalyst
            Call HandleAddCatalyst(UserIndex)
        Case ClientPacketID.RemoveCatalyst
            Call HandleRemoveCatalyst(UserIndex)
        Case ClientPacketID.CraftItem
            Call HandleCraftItem(UserIndex)
        Case ClientPacketID.CloseCrafting
            Call HandleCloseCrafting(UserIndex)
        Case ClientPacketID.MoveCraftItem
            Call HandleMoveCraftItem(UserIndex)
        Case ClientPacketID.PetLeaveAll
            Call HandlePetLeaveAll(UserIndex)
        Case ClientPacketID.GuardNoticeResponse
            Call HandleGuardNoticeResponse(UserIndex)
        Case ClientPacketID.GuardResendVerificationCode
            Call HandleGuardResendVerificationCode(UserIndex)
        Case ClientPacketID.ResetChar
            Call HandleResetChar(UserIndex)
        Case ClientPacketID.DeleteItem
            Call HandleDeleteItem(UserIndex)
        Case Else
            Err.raise -1, "Invalid Message"
    End Select
    
    If (Message.GetAvailable() > 0) Then
        Err.raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketID & "' se encuentra en mal estado con '" & Message.GetAvailable() & "' bytes de mas por el usuario '" & UserList(UserIndex).Name & "'"
    End If
    
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If Err.Number <> 0 Then
        Call RegistrarError(Err.Number, Err.Description & vbNewLine & "PackedID: " & PacketID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "UserName: " & UserList(UserIndex).Name, "UserIndex: " & UserIndex), "Protocol.HandleIncomingData", Erl)
        Call CloseSocket(UserIndex)
        
        HandleIncomingData = False
    End If
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

        On Error GoTo errHandler

        Dim UserName    As String
        Dim CuentaEmail As String
        Dim Password    As String
        Dim Version     As String
        Dim MD5         As String

102         CuentaEmail = Reader.ReadString8()
104         Password = Reader.ReadString8()
106         Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
108         UserName = Reader.ReadString8()
114         MD5 = Reader.ReadString8()

        #If DEBUGGING = False Then

116         If Not VersionOK(Version) Then
118             Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
120             Call CloseSocket(UserIndex)
                Exit Sub

            End If

        #End If
        
122     If EsGmChar(UserName) Then
            
124         If AdministratorAccounts(UCase$(UserName)) <> UCase$(CuentaEmail) Then
130             Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
        End If
  
132     If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MD5) Then
134         Call CloseSocket(UserIndex)
            Exit Sub

        End If

136     If Not AsciiValidos(UserName) Then
138         Call WriteShowMessageBox(UserIndex, "Nombre invalido.")
140         Call CloseSocket(UserIndex)
            Exit Sub

        End If

180     Call ConnectUser(UserIndex, UserName, CuentaEmail)

        Exit Sub
    
errHandler:
        
182     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)
184

End Sub

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo errHandler

        Dim UserName As String
        Dim race     As e_Raza
        Dim gender   As e_Genero
        Dim Hogar    As e_Ciudad
        Dim Class As e_Class
        Dim Head        As Integer
        Dim CuentaEmail As String
        Dim Password    As String
        Dim MD5         As String
        Dim Version     As String

102         CuentaEmail = Reader.ReadString8()
104         Password = Reader.ReadString8()
106         Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
108         UserName = Reader.ReadString8()
110         race = Reader.ReadInt8()
112         gender = Reader.ReadInt8()
114         Class = Reader.ReadInt8()
116         Head = Reader.ReadInt16()
118         Hogar = Reader.ReadInt8()
124         MD5 = Reader.ReadString8()

126     If PuedeCrearPersonajes = 0 Then
128         Call WriteShowMessageBox(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
130         Call CloseSocket(UserIndex)
            Exit Sub

        End If

132     If aClon.MaxPersonajes(UserList(UserIndex).IP) Then
134         Call WriteShowMessageBox(UserIndex, "Has creado demasiados personajes.")
136         Call CloseSocket(UserIndex)
            Exit Sub

        End If

        #If DEBUGGING = False Then

142         If Not VersionOK(Version) Then
144             Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
146             Call CloseSocket(UserIndex)
                Exit Sub

            End If

        #End If
        
148     If EsGmChar(UserName) Then
            
150         If AdministratorAccounts(UCase$(UserName)) <> UCase$(CuentaEmail) Then
152             Call WriteShowMessageBox(UserIndex, "El nombre de usuario ingresado está siendo ocupado por un miembro del Staff.")
154             Call CloseSocket(UserIndex)
                Exit Sub

            End If
            
        End If

156     If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MD5) Then
158         Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
160     If GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID) >= MAX_PERSONAJES Then
162         Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
164     If Not ConnectNewUser(UserIndex, UserName, race, gender, Class, Head, CuentaEmail, Hogar) Then
166         Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        Exit Sub
    
errHandler:

168     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
170

End Sub

Private Sub HandleThrowDice(ByVal UserIndex As Integer)
    
        On Error GoTo HandleThrowDice_Err

100     With UserList(UserIndex).Stats
102         .UserAtributos(e_Atributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
104         .UserAtributos(e_Atributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
106         .UserAtributos(e_Atributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
108         .UserAtributos(e_Atributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
110         .UserAtributos(e_Atributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)

        End With
    
112     Call WriteDiceRoll(UserIndex)

        Exit Sub

HandleThrowDice_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleThrowDice", Erl)
116
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()

            '[Consejeros & GMs]
104         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call LogGM(.Name, "Dijo: " & chat)
            End If
        
            'I see you....
108         If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
        
110             .flags.Oculto = 0
112             .Counters.TiempoOculto = 0
            
114             If .flags.Navegando = 1 Then

116                 If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
118                     Call EquiparBarco(UserIndex)
120                     Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", e_FontTypeNames.FONTTYPE_INFO)
122                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
124                     Call RefreshCharStatus(UserIndex)
                    End If

                Else

126                 If .flags.invisible = 0 Then
128                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", e_FontTypeNames.FONTTYPE_INFO)
130                     Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
    
                    End If

                End If

            End If
       
132         If .flags.Silenciado = 1 Then
        
                'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", e_FontTypeNames.FONTTYPE_VENENO)
134             Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
            
            Else

136             If LenB(chat) <> 0 Then
            
                    'Analize chat...
138                 Call Statistics.ParseChat(chat)

                    ' WyroX: Foto-denuncias - Push message
                    Dim i As Long
140                 For i = 1 To UBound(.flags.ChatHistory) - 1
142                     .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    
144                 .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                
146                 If .flags.Muerto = 1 Then
                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR, UserList(UserIndex).Name))
148                     Call SendData(SendTarget.ToUsuariosMuertos, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                    
                    Else
150                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor))

                    End If

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalk", Erl)
154

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
        
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If UserList(UserIndex).flags.Muerto = 1 Then
        
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", e_FontTypeNames.FONTTYPE_INFO)
            
            Else

                '[Consejeros & GMs]
108             If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
110                 Call LogGM(.Name, "Grito: " & chat)
                End If
            
                'I see you....
112             If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            
114                 .flags.Oculto = 0
116                 .Counters.TiempoOculto = 0
                
118                 If .flags.Navegando = 1 Then
                    
                        'TODO: Revisar con WyroX
120                     If .clase = e_Class.Pirat Then
                    
                            ' Pierde la apariencia de fragata fantasmal
122                         Call EquiparBarco(UserIndex)
124                         Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", e_FontTypeNames.FONTTYPE_INFO)
126                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
128                         Call RefreshCharStatus(UserIndex)
                        End If
    
                    Else
    
130                     If .flags.invisible = 0 Then
132                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
134                         Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", e_FontTypeNames.FONTTYPE_INFO)
    
                        End If
    
                    End If

                End If
            
136             If .flags.Silenciado = 1 Then
138                 Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        
                    'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", e_FontTypeNames.FONTTYPE_VENENO)
                Else

140                 If LenB(chat) <> 0 Then
                        'Analize chat...
142                     Call Statistics.ParseChat(chat)

                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Long
144                     For i = 1 To UBound(.flags.ChatHistory) - 1
146                         .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
                    
148                     .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat

150                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
               
                    End If

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:

152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleYell", Erl)
154

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat            As String
            Dim targetCharIndex As String
            Dim targetUserIndex As Integer

102         targetCharIndex = Reader.ReadString8()
104         chat = Reader.ReadString8()
    
106         If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(targetCharIndex)) < 0 Then Exit Sub
        
108         targetUserIndex = NameIndex(targetCharIndex)

110         If targetUserIndex <= 0 Then 'existe el usuario destino?
112             Call WriteConsoleMsg(UserIndex, "Usuario offline o inexistente.", e_FontTypeNames.FONTTYPE_INFO)

            Else

114             If EstaPCarea(UserIndex, targetUserIndex) Then

116                 If LenB(chat) <> 0 Then
                    
                        'Analize chat...
118                     Call Statistics.ParseChat(chat)

                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Long

120                     For i = 1 To UBound(.flags.ChatHistory) - 1
122                         .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
                        
124                     .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
126                     Call SendData(SendTarget.ToSuperioresArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, RGB(157, 226, 20)))
                        
128                     Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, RGB(157, 226, 20))
130                     Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, RGB(157, 226, 20))
                        'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_MP)
                        'Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_MP)
132                     Call WritePlayWave(targetUserIndex, e_FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                Else
134                 Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_MP)
136                 Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_MP)
138                 Call WritePlayWave(targetUserIndex, e_FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:

140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhisper", Erl)
142

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

        Dim Heading As e_Heading
    
100     With UserList(UserIndex)

102         Heading = Reader.ReadInt8()
        
104         If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then
        
106             If .flags.Comerciando Or .flags.Crafteando <> 0 Then Exit Sub

108             If .flags.Meditando Then
            
                    'Stop meditating, next action will start movement.
110                 .flags.Meditando = False
112                 UserList(UserIndex).Char.FX = 0
114                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))

                End If
                
                Dim CurrentTick As Long
116                 CurrentTick = GetTickCount
            
                'Prevent SpeedHack (refactored by WyroX)
118             If Not EsGM(UserIndex) Then
                    Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single
120                 ElapsedTimeStep = CurrentTick - .Counters.LastStep
122                 MinTimeStep = .Intervals.Caminar / .Char.speeding
124                 DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

126                 If DeltaStep > 0 Then
                
128                     .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
130                     If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                            'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración » Posible uso de SpeedHack del usuario " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
132                         Call WritePosUpdate(UserIndex)
                            Exit Sub

                        End If

                    Else
                
134                     .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

136                     If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                    End If

                End If
            
                'Move user
138             If MoveUserChar(UserIndex, Heading) Then
            
                    ' Save current step for anti-sh
140                 .Counters.LastStep = CurrentTick
                
142                 If UserList(UserIndex).Grupo.EnGrupo Then
144                     Call CompartirUbicacion(UserIndex)

                    End If
    
                    'Stop resting if needed
146                 If .flags.Descansar Then
148                     .flags.Descansar = False
                        
150                     Call WriteRestOK(UserIndex)
                        'Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", e_FontTypeNames.FONTTYPE_INFO)
152                     Call WriteLocaleMsg(UserIndex, "178", e_FontTypeNames.FONTTYPE_INFO)
    
                    End If
                        
154                 Call CancelExit(UserIndex)
                        
                    'Esta usando el /HOGAR, no se puede mover
156                 If .flags.Traveling = 1 Then
158                     .flags.Traveling = 0
160                     .Counters.goHome = 0
162                     Call WriteConsoleMsg(UserIndex, "Has cancelado el viaje a casa.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' Si no pudo moverse
                Else
164                 .Counters.LastStep = 0
166                 Call WritePosUpdate(UserIndex)

                End If

            Else    'paralized

168             If Not .flags.UltimoMensaje = 1 Then
170                 .flags.UltimoMensaje = 1
                    'Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", e_FontTypeNames.FONTTYPE_INFO)
172                 Call WriteLocaleMsg(UserIndex, "54", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If
            
            'Can't move while hidden except he is a thief
174         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                
176             If .clase <> e_Class.Thief And .clase <> e_Class.Bandit Then
            
178                 .flags.Oculto = 0
180                 .Counters.TiempoOculto = 0
                
182                 If .flags.Navegando = 1 Then
                        
184                     If .clase = e_Class.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
186                         Call EquiparBarco(UserIndex)
188                         Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", e_FontTypeNames.FONTTYPE_INFO)
190                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
192                         Call RefreshCharStatus(UserIndex)
                        End If
    
                    Else
    
                        'If not under a spell effect, show char
194                     If .flags.invisible = 0 Then
                            'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", e_FontTypeNames.FONTTYPE_INFO)
196                         Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
198                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                        End If
    
                    End If
    
                End If
                
            End If

        End With

        Exit Sub

HandleWalk_Err:
200     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWalk", Erl)
202
        
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

100     Call WritePosUpdate(UserIndex)
  
        Exit Sub

HandleRequestPositionUpdate_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandlRequestPositionUpdate", Erl)
104
        
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
        
100     With UserList(UserIndex)
    
            'If dead, can't attack
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡No podes atacar a nadie porque estas muerto!!.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If equiped weapon is ranged, can't attack this way
106         If .Invent.WeaponEqpObjIndex > 0 Then

108             If ObjData(.Invent.WeaponEqpObjIndex).Proyectil = 1 Then
110                 Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If
        
112         If .Invent.HerramientaEqpObjIndex > 0 Then
114             Call WriteConsoleMsg(UserIndex, "Para atacar debes desequipar la herramienta.", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If
        
116         If UserList(UserIndex).flags.Meditando Then
118             UserList(UserIndex).flags.Meditando = False
120             UserList(UserIndex).Char.FX = 0
122             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))

            End If
        
            'If exiting, cancel
124         Call CancelExit(UserIndex)
        
            'Attack!
126         Call UsuarioAtaca(UserIndex)
            
            'I see you...
128         If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
        
130             .flags.Oculto = 0
132             .Counters.TiempoOculto = 0
                
134             If .flags.Navegando = 1 Then

136                 If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
138                     Call EquiparBarco(UserIndex)
140                     Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", e_FontTypeNames.FONTTYPE_INFO)
142                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
144                     Call RefreshCharStatus(UserIndex)
                    End If
    
                Else
    
146                 If .flags.invisible = 0 Then
148                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", e_FontTypeNames.FONTTYPE_INFO)
150                     Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFOIAO)
    
                    End If
    
                End If
    
            End If

        End With

        Exit Sub

HandleAttack_Err:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAttack", Erl)
154
        
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
100     With UserList(UserIndex)

            'If dead, it can't pick up objects
102         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", e_FontTypeNames.FONTTYPE_INFO)
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Lower rank administrators can't pick up items
106         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
108             Call WriteConsoleMsg(UserIndex, "No podés tomar ningun objeto.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
110         Call PickObj(UserIndex)

        End With
        
        Exit Sub

HandlePickUp_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePickUp", Erl)
114
        
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
100     With UserList(UserIndex)
            
            If esCiudadano(UserIndex) Then
102             If .flags.Seguro Then
104                 Call WriteSafeModeOff(UserIndex)
                Else
106                 Call WriteSafeModeOn(UserIndex)
                End If
                
108             .flags.Seguro = Not .flags.Seguro
            Else
                Call WriteConsoleMsg(UserIndex, "Solo los ciudadanos pueden cambiar el seguro", e_FontTypeNames.FONTTYPE_TALK)
            End If

        End With

        Exit Sub

HandleSafeToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSafeToggle", Erl)
112
        
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
100     With UserList(UserIndex)
        
102         .flags.SeguroParty = Not .flags.SeguroParty
        
104         If .flags.SeguroParty Then
106             Call WritePartySafeOn(UserIndex)
            
            Else
108             Call WritePartySafeOff(UserIndex)

            End If

        End With

        Exit Sub

HandlePartyToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePartyToggle", Erl)
112
        
End Sub

Private Sub HandleSeguroClan(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSeguroClan_Err

        '***************************************************
        'Author: Ladder
        'Date: 31/10/20
        '***************************************************
100     With UserList(UserIndex)

102         .flags.SeguroClan = Not .flags.SeguroClan

104         Call WriteClanSeguro(UserIndex, .flags.SeguroClan)

        End With

        Exit Sub

HandleSeguroClan_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSeguroClan", Erl)
108
        
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

100     Call modGuilds.SendGuildLeaderInfo(UserIndex)

        Exit Sub

HandleRequestGuildLeaderInfo_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
104
        
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

100     Call WriteAttributes(UserIndex)

        Exit Sub

HandleRequestAtributes_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAtributes", Erl)
104
        
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

100     Call WriteSendSkills(UserIndex)

        Exit Sub

HandleRequestSkills_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestSkills", Erl)
104
        
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

100     Call WriteMiniStats(UserIndex)

        Exit Sub

HandleRequestMiniStats_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMiniStats", Erl)
104
        
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
100     If UserList(UserIndex).flags.TargetNPC <> 0 Then
    
102         If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
104             Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

            End If

        End If

106     UserList(UserIndex).flags.Comerciando = False

108     Call WriteCommerceEnd(UserIndex)
 
        Exit Sub

HandleCommerceEnd_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
112
        
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
100     With UserList(UserIndex)

        
        
        
            'Quits commerce mode with user
102         If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
104             Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", e_FontTypeNames.FONTTYPE_TALK)
106             Call FinComerciarUsu(.ComUsu.DestUsu)
            
                'Send data in the outgoing buffer of the other user

            End If
        
108         Call FinComerciarUsu(UserIndex)

        End With
        
        Exit Sub

HandleUserCommerceEnd_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
112
        
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
100     With UserList(UserIndex)

            'User exits banking mode
102         .flags.Comerciando = False
        
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("171", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
106         Call WriteBankEnd(UserIndex)

        End With
        
        Exit Sub

HandleBankEnd_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
110
        
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
100     Call AceptarComercioUsu(UserIndex)
        
        Exit Sub

HandleUserCommerceOk_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOk", Erl)
104
        
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
    
100     With UserList(UserIndex)

102         otherUser = .ComUsu.DestUsu
        
            'Offer rejected
104         If otherUser > 0 Then
106             If UserList(otherUser).flags.UserLogged Then
108                 Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", e_FontTypeNames.FONTTYPE_TALK)
110                 Call FinComerciarUsu(otherUser)
                
                    'Send data in the outgoing buffer of the other user

                End If

            End If
        
112         Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", e_FontTypeNames.FONTTYPE_TALK)
114         Call FinComerciarUsu(UserIndex)

        End With
        
        Exit Sub

HandleUserCommerceReject_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceReject", Erl)
118
        
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

        Dim Slot   As Byte
        Dim amount As Long
    
100     With UserList(UserIndex)

102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt32()

106         If Not IntervaloPermiteTirar(UserIndex) Then Exit Sub

108         If amount <= 0 Then Exit Sub

            'low rank admins can't drop item. Neither can the dead nor those sailing or riding a horse.
110         If .flags.Muerto = 1 Then Exit Sub
                      
            'If the user is trading, he can't drop items => He's cheating, we kick him.
112         If .flags.Comerciando Then Exit Sub
    
            'Si esta navegando y no es pirata, no dejamos tirar items al agua.
114         If .flags.Navegando = 1 And Not .clase = e_Class.Pirat Then
116             Call WriteConsoleMsg(UserIndex, "Solo los Piratas pueden tirar items en altamar", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
118         If .flags.Montado = 1 Then
120             Call WriteConsoleMsg(UserIndex, "Debes descender de tu montura para dejar objetos en el suelo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Are we dropping gold or other items??
122         If Slot = FLAGORO Then
124             Call TirarOro(amount, UserIndex)
            
            Else
        
                '04-05-08 Ladder
126             If (.flags.Privilegios And e_PlayerType.Admin) <> 16 Then
128                 If EsNewbie(UserIndex) And ObjData(.Invent.Object(Slot).ObjIndex).Newbie = 1 Then
130                     Call WriteConsoleMsg(UserIndex, "No se pueden tirar los objetos Newbies.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
            
132                 If ObjData(.Invent.Object(Slot).ObjIndex).Intirable = 1 And Not EsGM(UserIndex) Then
134                     Call WriteConsoleMsg(UserIndex, "Acción no permitida.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
136                 ElseIf ObjData(.Invent.Object(Slot).ObjIndex).Intirable = 1 And EsGM(UserIndex) Then
138                     If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
140                         If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
142                         Call DropObj(UserIndex, Slot, amount, .Pos.Map, .Pos.x, .Pos.Y)
                        End If
                        Exit Sub
                    End If
                
144                 If ObjData(.Invent.Object(Slot).ObjIndex).Instransferible = 1 Then
146                     Call WriteConsoleMsg(UserIndex, "Acción no permitida.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
            

                End If
        
148             If ObjData(.Invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otBarcos And UserList(UserIndex).flags.Navegando Then
150                 Call WriteConsoleMsg(UserIndex, "Para tirar la barca deberias estar en tierra firme.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
                '04-05-08 Ladder
        
                'Only drop valid slots
152             If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
            
154                 If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

156                 Call DropObj(UserIndex, Slot, amount, .Pos.Map, .Pos.x, .Pos.Y)

                End If

            End If

        End With
        
        Exit Sub

HandleDrop_Err:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDrop", Erl)
160
        
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

100     With UserList(UserIndex)

            Dim Spell As Byte
102             Spell = Reader.ReadInt8()
        
104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", e_FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         .flags.Hechizo = Spell
        
110         If .flags.Hechizo < 1 Or .flags.Hechizo > MAXUSERHECHIZOS Then
112             .flags.Hechizo = 0
            End If
        
114         If .flags.Hechizo <> 0 Then

116             If (.flags.Privilegios And e_PlayerType.Consejero) = 0 Then
                    
                    If .Stats.UserHechizos(Spell) <> 0 Then
                    
120                     If Hechizos(.Stats.UserHechizos(Spell)).AutoLanzar = 1 Then
122                         UserList(UserIndex).flags.TargetUser = UserIndex
124                         Call LanzarHechizo(.flags.Hechizo, UserIndex)
    
                        Else
126                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Magia)
        
                        End If
                    
                    End If
                    
                End If

            End If
        
        End With
        
        Exit Sub

HandleCastSpell_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCastSpell", Erl)
130
        
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

100     With UserList(UserIndex)

            Dim x As Byte
            Dim Y As Byte
        
102         x = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
        
106         Call LookatTile(UserIndex, .Pos.Map, x, Y)

        End With

        Exit Sub

HandleLeftClick_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeftClick", Erl)
110
        
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

100     With UserList(UserIndex)

            Dim x As Byte
            Dim Y As Byte
        
102         x = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
        
106         Call Accion(UserIndex, .Pos.Map, x, Y)

        End With
        
        Exit Sub

HandleDoubleClick_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoubleClick", Erl)
110
        
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

100     With UserList(UserIndex)

            Dim Skill As e_Skill
102             Skill = Reader.ReadInt8()
        
104         If UserList(UserIndex).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", e_FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If exiting, cancel
108         Call CancelExit(UserIndex)
        
110         Select Case Skill

                Case Robar, Magia, Domar
112                 Call WriteWorkRequestTarget(UserIndex, Skill)

114             Case Ocultarse

116                 If .flags.Montado = 1 Then

                        '[CDT 17-02-2004]
118                     If Not .flags.UltimoMensaje = 3 Then
120                         Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás montado.", e_FontTypeNames.FONTTYPE_INFO)
122                         .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

124                 If .flags.Oculto = 1 Then

                        '[CDT 17-02-2004]
126                     If Not .flags.UltimoMensaje = 2 Then
128                         Call WriteLocaleMsg(UserIndex, "55", e_FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", e_FontTypeNames.FONTTYPE_INFO)
130                         .flags.UltimoMensaje = 2

                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                    
132                 If .flags.EnReto Then
134                     Call WriteConsoleMsg(UserIndex, "No podés ocultarte durante un reto.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
136                 If .flags.EnConsulta Then
138                     Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estas en consulta.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                    
140                 If MapInfo(.Pos.Map).SinInviOcul Then
142                     Call WriteConsoleMsg(UserIndex, "Una fuerza divina te impide ocultarte en esta zona.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
144                 Call DoOcultarse(UserIndex)

            End Select

        End With
        
        Exit Sub

HandleWork_Err:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWork", Erl)
148
        
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
100     With UserList(UserIndex)

102         Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos", e_FontTypeNames.FONTTYPE_VENENO))
104         Call WriteShowMessageBox(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        
106         Call CloseSocket(UserIndex)

        End With
        
        Exit Sub

HandleUseSpellMacro_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseSpellMacro", Erl)
110
        
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
    
100     With UserList(UserIndex)

            Dim Slot As Byte
102             Slot = Reader.ReadInt8()
        
104         If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
106             If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

108             Call UseInvItem(UserIndex, Slot)
                
            End If

        End With

        Exit Sub

HandleUseItem_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseItem", Erl)
112
        
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

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item < 1 Then Exit Sub
        
            ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
106         Call HerreroConstruirItem(UserIndex, Item)

        Exit Sub

HandleCraftBlacksmith_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftBlacksmith", Erl)
110
        
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

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item = 0 Then Exit Sub

106         Call CarpinteroConstruirItem(UserIndex, Item)

        Exit Sub

HandleCraftCarpenter_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftCarpenter", Erl)
110
        
End Sub

Private Sub HandleCraftAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftAlquimia_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item < 1 Then Exit Sub

        Exit Sub

HandleCraftAlquimia_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftAlquimia", Erl)
108
        
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftSastre_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item < 1 Then Exit Sub

106         Call SastreConstruirItem(UserIndex, Item)

        Exit Sub

HandleCraftSastre_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftSastre", Erl)
110
        
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

100     With UserList(UserIndex)
        
            Dim x        As Byte
            Dim Y        As Byte

            Dim Skill    As e_Skill
            Dim DummyInt As Integer

            Dim tU       As Integer   'Target user
            Dim tN       As Integer   'Target NPC
        
102         x = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
            
106         Skill = Reader.ReadInt8()

            .Trabajo.Target_X = x
            .Trabajo.Target_Y = Y
            .Trabajo.TargetSkill = Skill
            
108         If .flags.Muerto = 1 Or .flags.Descansar Or Not InMapBounds(.Pos.Map, x, Y) Then Exit Sub

110         If Not InRangoVision(UserIndex, x, Y) Then
112             Call WritePosUpdate(UserIndex)
                Exit Sub

            End If
            
114         If .flags.Meditando Then
116             .flags.Meditando = False
118             .Char.FX = 0
120             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))

            End If
        
            'If exiting, cancel
122         Call CancelExit(UserIndex)
            
124         Select Case Skill

                    Dim consumirMunicion As Boolean

                Case e_Skill.Proyectiles
            
                    'Check attack interval
126                 If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                    'Check Magic interval
128                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                    'Check bow's interval
130                 If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                    'Make sure the item is valid and there is ammo equipped.
132                 With .Invent

134                     If .WeaponEqpObjIndex = 0 Then
136                         DummyInt = 1
138                     ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
140                         DummyInt = 1
142                     ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
144                         DummyInt = 1
146                     ElseIf .MunicionEqpObjIndex = 0 Then
148                         DummyInt = 1
150                     ElseIf ObjData(.WeaponEqpObjIndex).Proyectil <> 1 Then
152                         DummyInt = 2
154                     ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> e_OBJType.otFlechas Then
156                         DummyInt = 1
158                     ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
160                         DummyInt = 1

                        End If
                    
162                     If DummyInt <> 0 Then
164                         If DummyInt = 1 Then
166                             Call WriteConsoleMsg(UserIndex, "No tenés municiones.", e_FontTypeNames.FONTTYPE_INFO)

                            End If
                        
168                         Call Desequipar(UserIndex, .MunicionEqpSlot)
170                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If

                    End With
                
                    'Quitamos stamina
172                 If .Stats.MinSta >= 10 Then
174                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                    Else
180                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "Estís muy cansado para luchar.", e_FontTypeNames.FONTTYPE_INFO)
182                     Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub

                    End If
                
184                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
186                 tU = .flags.TargetUser
188                 tN = .flags.TargetNPC
190                 consumirMunicion = False

                    'Validate target
192                 If tU > 0 Then

                        'Only allow to atack if the other one can retaliate (can see us)
194                     If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
196                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para atacar.", e_FontTypeNames.FONTTYPE_WARNING)
198                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If
                    
                        'Prevent from hitting self
200                     If tU = UserIndex Then
202                         Call WriteConsoleMsg(UserIndex, "¡No podés atacarte a vos mismo!", e_FontTypeNames.FONTTYPE_INFO)
204                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If
                    
                        'Attack!
206                     If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    
                        Dim backup    As Byte
                        Dim envie     As Boolean
                        Dim Particula As Integer
                        Dim Tiempo    As Long

                        ' Porque no es HandleAttack ???
208                     Call UsuarioAtacaUsuario(UserIndex, tU)

210                     If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
212                         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateFX(UserList(tU).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                        End If
                        
                        'Si no es GM invisible, le envio el movimiento del arma.
                        If UserList(UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
                        End If
                    
214                     If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
                    
216                         Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
218                         Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
220                         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, Particula, Tiempo, False))

                        End If
                    
222                     consumirMunicion = True
                    
224                 ElseIf tN > 0 Then

                        'Only allow to atack if the other one can retaliate (can see us)
226                     If Abs(NpcList(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(NpcList(tN).Pos.x - .Pos.x) > RANGO_VISION_X Then
228                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
230                         Call WriteWorkRequestTarget(UserIndex, 0)
                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", e_FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub

                        End If
                    
                        'Is it attackable???
232                     If NpcList(tN).Attackable <> 0 Then
234                         If PuedeAtacarNPC(UserIndex, tN) Then
236                             Call UsuarioAtacaNpc(UserIndex, tN)
238                             consumirMunicion = True
                                'Si no es GM invisible, le envio el movimiento del arma.
                                If UserList(UserIndex).flags.AdminInvisible = 0 Then
                                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
                                End If
                            Else
240                             consumirMunicion = False

                            End If

                        End If

                    End If
                
242                 With .Invent
244                     DummyInt = .MunicionEqpSlot
                        
                        If DummyInt <> 0 Then
                        
                            'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
246                         If consumirMunicion Then
248                             Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            End If
                        
250                         If .Object(DummyInt).amount > 0 Then

                                'QuitarUserInvItem unequipps the ammo, so we equip it again
252                             .MunicionEqpSlot = DummyInt
254                             .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
256                             .Object(DummyInt).Equipped = 1
    
                            Else
258                             .MunicionEqpSlot = 0
260                             .MunicionEqpObjIndex = 0
    
                            End If
    
262                         Call UpdateUserInv(False, UserIndex, DummyInt)
                        
                        End If
                        
                    End With
                    '-----------------------------------
            
264             Case e_Skill.Magia
                    'Check the map allows spells to be casted.
                    '  If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    ' Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", e_FontTypeNames.FONTTYPE_FIGHT)
                    '  Exit Sub
                    ' End If
                
                    'Target whatever is in that tile
266                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
                    'If it's outside range log it and exit
268                 If Abs(.Pos.x - x) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
270                     Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.x & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & x & "/" & Y & ")")
                        Exit Sub

                    End If
                
                    'Check bow's interval
272                 If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                    'Check attack-spell interval
274                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                    'Check Magic interval
276                 If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                
                    'Check intervals and cast
278                 If .flags.Hechizo > 0 Then
280                     Call LanzarHechizo(.flags.Hechizo, UserIndex)
282                     .flags.Hechizo = 0
                    Else
284                     Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", e_FontTypeNames.FONTTYPE_INFO)

                    End If
            
286             Case e_Skill.Pescar
                    If .Counters.Trabajando = 0 Then
                        Call Trabajar(UserIndex, e_Skill.Pescar)
                    End If
348             Case e_Skill.Talar
                    If .Counters.Trabajando = 0 Then
                        Call Trabajar(UserIndex, e_Skill.Talar)
                    End If
                    
400             Case e_Skill.Alquimia
            
402                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
404                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
                    
                    'Check interval
406                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

408                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 3  ' Herramientas de Alquimia - Tijeras

410                         If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
412                             Call WriteWorkRequestTarget(UserIndex, 0)
414                             Call WriteConsoleMsg(UserIndex, "Esta prohibido cortar raices en las ciudades.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                            
416                         If MapData(.Pos.Map, x, Y).ObjInfo.amount <= 0 Then
418                             Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas raices.", e_FontTypeNames.FONTTYPE_INFO)
420                             Call WriteWorkRequestTarget(UserIndex, 0)
422                             Call WriteMacroTrabajoToggle(UserIndex, False)
                                Exit Sub

                            End If
                
424                         DummyInt = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex
                            
426                         If DummyInt > 0 Then
                            
428                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
430                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
432                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
434                             If .Pos.x = x And .Pos.Y = Y Then
436                                 Call WriteConsoleMsg(UserIndex, "No podés quitar raices allí.", e_FontTypeNames.FONTTYPE_INFO)
438                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
                                '¡Hay un arbol donde clickeo?
440                             If ObjData(DummyInt).OBJType = e_OBJType.otArboles Then
442                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TIJERAS, .Pos.x, .Pos.Y))
444                                 Call DoRaices(UserIndex, x, Y)

                                End If

                            Else
446                             Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", e_FontTypeNames.FONTTYPE_INFO)
448                             Call WriteWorkRequestTarget(UserIndex, 0)
450                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If
                
                    End Select
                
452             Case e_Skill.Mineria
                    If .Counters.Trabajando = 0 Then
                        Call Trabajar(UserIndex, e_Skill.Mineria)
                    End If
500             Case e_Skill.Robar

                    'Does the map allow us to steal here?
502                 If MapInfo(.Pos.Map).Seguro = 0 Then
                    
                        'Check interval
504                     If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                    
                        'Target whatever is in that tile
506                     Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
508                     tU = .flags.TargetUser
                    
510                     If tU > 0 And tU <> UserIndex Then

                            'Can't steal administrative players
512                         If UserList(tU).flags.Privilegios And e_PlayerType.user Then
514                             If UserList(tU).flags.Muerto = 0 Then
                                    Dim DistanciaMaxima As Integer

516                                 If .clase = e_Class.Thief Then
518                                     DistanciaMaxima = 2
                                    Else
520                                     DistanciaMaxima = 1

                                    End If

522                                 If Abs(.Pos.x - UserList(tU).Pos.x) + Abs(.Pos.Y - UserList(tU).Pos.Y) > DistanciaMaxima Then
524                                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                        'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
526                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
                                    '17/09/02
                                    'Check the trigger
528                                 If MapData(UserList(tU).Pos.Map, UserList(tU).Pos.x, UserList(tU).Pos.Y).trigger = e_Trigger.ZonaSegura Then
530                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", e_FontTypeNames.FONTTYPE_WARNING)
532                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
534                                 If MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = e_Trigger.ZonaSegura Then
536                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", e_FontTypeNames.FONTTYPE_WARNING)
538                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
540                                 Call DoRobar(UserIndex, tU)

                                End If

                            End If

                        Else
542                         Call WriteConsoleMsg(UserIndex, "No a quien robarle!", e_FontTypeNames.FONTTYPE_INFO)
544                         Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else
546                     Call WriteConsoleMsg(UserIndex, "¡No podés robar en zonas seguras!", e_FontTypeNames.FONTTYPE_INFO)
548                     Call WriteWorkRequestTarget(UserIndex, 0)

                    End If
                    
550             Case e_Skill.Domar
                    'Modificado 25/11/02
                    'Optimizado y solucionado el bug de la doma de criaturas hostiles.
                    
                    'Target whatever is that tile
552                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
554                 tN = .flags.TargetNPC
                    
556                 If tN > 0 Then
558                     If NpcList(tN).flags.Domable > 0 Then
560                         If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 4 Then
562                             Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                            
564                         If LenB(NpcList(tN).flags.AttackedBy) <> 0 Then
566                             Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que esta luchando con un jugador.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                            
568                         Call DoDomar(UserIndex, tN)
                        Else
570                         Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", e_FontTypeNames.FONTTYPE_INFO)
    
                        End If
    
                    Else
572                     Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!", e_FontTypeNames.FONTTYPE_INFO)
    
                    End If
               
574             Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                    'Check interval
576                 If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                
578                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
                    'Check there is a proper item there
580                 If .flags.TargetObj > 0 Then
582                     If ObjData(.flags.TargetObj).OBJType = e_OBJType.otFragua Then

                            'Validate other items
584                         If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                                Exit Sub

                            End If
                        
                            ''chequeamos que no se zarpe duplicando oro
586                         If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
588                             If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
590                                 Call WriteConsoleMsg(UserIndex, "No tienes más minerales", e_FontTypeNames.FONTTYPE_INFO)
592                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                            
                                ''FUISTE
594                             Call WriteShowMessageBox(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            
596                             Call CloseSocket(UserIndex)
                                Exit Sub

                            End If
                        
598                         Call FundirMineral(UserIndex)
                        
                        Else
                    
600                         Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", e_FontTypeNames.FONTTYPE_INFO)
602                         Call WriteWorkRequestTarget(UserIndex, 0)

604                         If UserList(UserIndex).Counters.Trabajando > 1 Then
606                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If

                        End If

                    Else
                
608                     Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", e_FontTypeNames.FONTTYPE_INFO)
610                     Call WriteWorkRequestTarget(UserIndex, 0)

612                     If UserList(UserIndex).Counters.Trabajando > 1 Then
614                         Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If

                    End If

616             Case e_Skill.Grupo
                    'If UserList(UserIndex).Grupo.EnGrupo = False Then
                    'Target whatever is in that tile
                    'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
618                 tU = .flags.TargetUser
                    
                    'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
620                 If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
622                     If UserList(UserIndex).Grupo.EnGrupo = False Then
624                         If UserList(tU).flags.Muerto = 0 Then
626                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 8 Then
628                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
630                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                         
632                             If UserList(UserIndex).Grupo.CantidadMiembros = 0 Then
634                                 UserList(UserIndex).Grupo.Lider = UserIndex
636                                 UserList(UserIndex).Grupo.Miembros(1) = UserIndex
638                                 UserList(UserIndex).Grupo.CantidadMiembros = 1
640                                 Call InvitarMiembro(UserIndex, tU)
                                Else
642                                 UserList(UserIndex).Grupo.Lider = UserIndex
644                                 Call InvitarMiembro(UserIndex, tU)

                                End If
                                         
                            Else
646                             Call WriteLocaleMsg(UserIndex, "7", e_FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", e_FontTypeNames.FONTTYPE_INFOIAO)
648                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                        Else

650                         If UserList(UserIndex).Grupo.Lider = UserIndex Then
652                             Call InvitarMiembro(UserIndex, tU)
                            Else
654                             Call WriteConsoleMsg(UserIndex, "Tu no podés invitar usuarios, debe hacerlo " & UserList(UserList(UserIndex).Grupo.Lider).Name & ".", e_FontTypeNames.FONTTYPE_INFOIAO)
656                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                        End If

                    Else
658                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' End If
660             Case e_Skill.MarcaDeClan

                    'If UserList(UserIndex).Grupo.EnGrupo = False Then
                    'Target whatever is in that tile
                    Dim clan_nivel As Byte
                
662                 If UserList(UserIndex).GuildIndex = 0 Then
664                     Call WriteConsoleMsg(UserIndex, "Servidor » No perteneces a ningún clan.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                
666                 clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

668                 If clan_nivel < 3 Then
670                     Call WriteConsoleMsg(UserIndex, "Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                                
672                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
674                 tU = .flags.TargetUser

676                 If tU = 0 Then Exit Sub
                    
678                 If UserList(UserIndex).GuildIndex = UserList(tU).GuildIndex Then
680                     Call WriteConsoleMsg(UserIndex, "Servidor » No podes marcar a un miembro de tu clan.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
                    
                    'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
682                 If tU > 0 And tU <> UserIndex Then

                        ' WyroX: No puede marcar admins invisibles
684                     If UserList(tU).flags.AdminInvisible <> 0 Then Exit Sub

                        'Can't steal administrative players
686                     If UserList(tU).flags.Muerto = 0 Then

                            'call marcar
688                         If UserList(tU).flags.invisible = 1 Or UserList(tU).flags.Oculto = 1 Then
690                             Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 50, False))
                            Else
692                             Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 150, False))

                            End If

694                         Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageConsoleMsg("Clan> [" & UserList(UserIndex).Name & "] marco a " & UserList(tU).Name & ".", e_FontTypeNames.FONTTYPE_GUILD))
                        Else
696                         Call WriteLocaleMsg(UserIndex, "7", e_FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", e_FontTypeNames.FONTTYPE_INFOIAO)
698                         Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else
700                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)

                    End If

702             Case e_Skill.MarcaDeGM
704                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
706                 tU = .flags.TargetUser

708                 If tU > 0 Then
710                     Call WriteConsoleMsg(UserIndex, "Servidor » [" & UserList(tU).Name & "] seleccionado.", e_FontTypeNames.FONTTYPE_SERVER)
                    Else
712                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                    
            End Select

        End With
        
        Exit Sub

HandleWorkLeftClick_Err:
714     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorkLeftClick", Erl)
716
        
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

    On Error GoTo errHandler

100 With UserList(UserIndex)
        
        Dim Desc       As String
        Dim GuildName  As String
        Dim errorStr   As String
        Dim Alineacion As Byte
        
102     Desc = Reader.ReadString8()
104     GuildName = Reader.ReadString8()
106     Alineacion = Reader.ReadInt8()
        
108     If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Alineacion, errorStr) Then

110         Call QuitarObjetos(407, 1, UserIndex)
112         Call QuitarObjetos(408, 1, UserIndex)
114         Call QuitarObjetos(409, 1, UserIndex)
116         Call QuitarObjetos(411, 1, UserIndex)
            
            
                
118             Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & " ha fundado el clan <" & GuildName & "> de alineación " & GuildAlignment(.GuildIndex) & ".", e_FontTypeNames.FONTTYPE_GUILD))
120             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
                'Update tag
122             Call RefreshCharStatus(UserIndex)
            Else
124             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
errHandler:

126 Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNewGuild", Erl)
128

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

100     With UserList(UserIndex)
        
            Dim spellSlot As Byte
            Dim Spell     As Integer
        
102         spellSlot = Reader.ReadInt8()
        
            'Validate slot
104         If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
106             Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate spell in the slot
108         Spell = .Stats.UserHechizos(spellSlot)

110         If Spell > 0 And Spell < NumeroHechizos + 1 Then

112             With Hechizos(Spell)
                    'Send information
114                 Call WriteConsoleMsg(UserIndex, "HECINF*" & Spell, e_FontTypeNames.FONTTYPE_INFO)

                End With

            End If

        End With
        
        Exit Sub

HandleSpellInfo_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpellInfo", Erl)
118
        
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
    
100     With UserList(UserIndex)
        
            Dim itemSlot As Byte
102             itemSlot = Reader.ReadInt8()
        
            'Dead users can't equip items
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate item slot
108         If itemSlot > UserList(UserIndex).CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
110         If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
112         Call EquiparInvItem(UserIndex, itemSlot)

        End With
        
        Exit Sub

HandleEquipItem_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEquipItem", Erl)
116
        
End Sub

''
' Handles the "Change_Heading" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChange_Heading(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChange_Heading_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Heading As e_Heading
102             Heading = Reader.ReadInt8()
        
            'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
104         If Heading > 0 And Heading < 5 Then
106             .Char.Heading = Heading
108             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

            End If

        End With

        Exit Sub

HandleChange_Heading_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChange_Heading", Erl)
112
        
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

100     With UserList(UserIndex)

            Dim i                      As Long
            Dim Count                  As Integer
            Dim points(1 To NUMSKILLS) As Byte
        
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
102         For i = 1 To NUMSKILLS
104             points(i) = Reader.ReadInt8()
            
106             If points(i) < 0 Then
108                 Call LogHackAttemp(.Name & " IP:" & .IP & " trató de hackear los skills.")
110                 .Stats.SkillPts = 0
112                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If
            
114             Count = Count + points(i)
116         Next i
        
118         If Count > .Stats.SkillPts Then
120             Call LogHackAttemp(.Name & " IP:" & .IP & " trató de hackear los skills.")
122             Call CloseSocket(UserIndex)
                Exit Sub

            End If

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
124         With .Stats

126             For i = 1 To NUMSKILLS
128                 .SkillPts = .SkillPts - points(i)
                    
                    If .UserSkills(i) <> .UserSkills(i) + points(i) Then
130                     .UserSkills(i) = .UserSkills(i) + points(i)
                    
                        'Client should prevent this, but just in case...
132                     If .UserSkills(i) > 100 Then
134                         .SkillPts = .SkillPts + .UserSkills(i) - 100
136                         .UserSkills(i) = 100
                        End If
                        
                        UserList(UserIndex).flags.ModificoSkills = True
                    End If
138             Next i

            End With

        End With
        
        Exit Sub

HandleModifySkills_Err:
140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleModifySkills", Erl)
142
        
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
    
100     With UserList(UserIndex)
        
            Dim SpawnedNpc As Integer
            Dim PetIndex   As Byte
        
102         PetIndex = Reader.ReadInt8()
        
104         If .flags.TargetNPC = 0 Then Exit Sub
        
106         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Entrenador Then Exit Sub
        
108         If NpcList(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
        
110             If PetIndex > 0 And PetIndex < NpcList(.flags.TargetNPC).NroCriaturas + 1 Then
                    'Create the creature
112                 SpawnedNpc = SpawnNpc(NpcList(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, NpcList(.flags.TargetNPC).Pos, True, False)
                
114                 If SpawnedNpc > 0 Then
116                     NpcList(SpawnedNpc).MaestroNPC = .flags.TargetNPC
118                     NpcList(.flags.TargetNPC).Mascotas = NpcList(.flags.TargetNPC).Mascotas + 1

                    End If

                End If

            Else
120             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))

            End If

        End With
        
        Exit Sub

HandleTrain_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrain", Erl)
124
        
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
    
100     With UserList(UserIndex)
        
            Dim Slot   As Byte
            Dim amount As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
        
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'El target es un NPC valido?
110         If .flags.TargetNPC < 1 Then Exit Sub
            
            'íEl NPC puede comerciar?
112         If NpcList(.flags.TargetNPC).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'Only if in commerce mode....
116         If Not .flags.Comerciando Then
118             Call WriteConsoleMsg(UserIndex, "No estás comerciando", e_FontTypeNames.FONTTYPE_INFO)
120             Call WriteCommerceEnd(UserIndex)
                Exit Sub

            End If
        
            'User compra el item
122         Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, amount)

        End With

        Exit Sub

HandleCommerceBuy_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceBuy", Erl)
126
        
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

100     With UserList(UserIndex)

            Dim Slot        As Byte
            Dim slotdestino As Byte
            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce
108         If .flags.Muerto = 1 Then
110             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            '¿El target es un NPC valido?
112         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿Es el banquero?
114         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then Exit Sub

            'User retira el item del slot
116         Call UserRetiraItem(UserIndex, Slot, amount, slotdestino)

        End With

        Exit Sub

HandleBankExtractItem_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractItem", Erl)
120
        
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
    
100     With UserList(UserIndex)

            Dim Slot   As Byte
            Dim amount As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
        
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
110         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
112         If NpcList(.flags.TargetNPC).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'User compra el item del slot
116         Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, amount)

        End With

        Exit Sub

HandleCommerceSell_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSell", Erl)
120
        
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
    
100     With UserList(UserIndex)
        
            Dim Slot        As Byte
            Dim slotdestino As Byte
            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce...
108         If .flags.Muerto = 1 Then
110             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
112         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
114         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then
                Exit Sub

            End If
            
116         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'User deposita el item del slot rdata
120         Call UserDepositaItem(UserIndex, Slot, amount, slotdestino)

        End With
        
        Exit Sub

HandleBankDeposit_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDeposit", Erl)
124
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim File     As String
            Dim title    As String
            Dim Msg      As String
            Dim postFile As String
            Dim handle   As Integer
            Dim i        As Long
            Dim Count    As Integer
        
102         title = Reader.ReadString8()
104         Msg = Reader.ReadString8()
        
106         If .flags.TargetObj > 0 Then
108             File = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
110             If FileExist(File, vbNormal) Then
112                 Count = val(GetVar(File, "INFO", "CantMSG"))
                
                    'If there are too many messages, delete the forum
114                 If Count > MAX_MENSAJES_FORO Then

116                     For i = 1 To Count
118                         Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
120                     Next i

122                     Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
124                     Count = 0

                    End If

                Else
                    'Starting the forum....
126                 Count = 0

                End If
            
128             handle = FreeFile()
130             postFile = Left$(File, Len(File) - 4) & CStr(Count + 1) & ".for"
            
                'Create file
132             Open postFile For Output As handle
134             Print #handle, title
136             Print #handle, Msg
138             Close #handle
            
                'Update post count
140             Call WriteVar(File, "INFO", "CantMSG", Count + 1)

            End If

        End With
        
        Exit Sub
        
errHandler:
142     Close #handle
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForumPost", Erl)
146

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

            Dim dir As Integer
        
102         If Reader.ReadBool() Then
104             dir = 1
            Else
106             dir = -1

            End If
        
108         Call DesplazarHechizo(UserIndex, dir, Reader.ReadInt8())

        Exit Sub

HandleMoveSpell_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
112
        
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
        
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Desc As String
        
102         Desc = Reader.ReadString8()
        
104         Call modGuilds.ChangeCodexAndDesc(Desc, .GuildIndex)

        End With
        
        Exit Sub
        
errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
108

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
    
100     With UserList(UserIndex)

            Dim tUser  As Integer
            Dim Slot   As Byte
            Dim amount As Long
            
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt32()
        
            'Get the other player
106         tUser = .ComUsu.DestUsu
        
            'If Amount is invalid, or slot is invalid and it's not gold, then ignore it.
108         If ((Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Or amount <= 0 Then Exit Sub
        
            'Is the other player valid??
110         If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
            'Is the commerce attempt valid??
112         If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
114             Call FinComerciarUsu(UserIndex)
                Exit Sub

            End If
        
            'Is he still logged??
116         If Not UserList(tUser).flags.UserLogged Then
118             Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else

                'Is he alive??
120             If UserList(tUser).flags.Muerto = 1 Then
122                 Call FinComerciarUsu(UserIndex)
                    Exit Sub

                End If
            
                'Has he got enough??
124             If Slot = FLAGORO Then

                    'gold
126                 If amount > .Stats.GLD Then
128                     Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                Else

                    'inventory
130                 If amount > .Invent.Object(Slot).amount Then
132                     Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If
                
134                 If .Invent.Object(Slot).ObjIndex > 0 Then
136                     If ObjData(.Invent.Object(Slot).ObjIndex).Instransferible = 1 Then
138                         Call WriteConsoleMsg(UserIndex, "Este objeto es intransferible, no podés venderlo.", e_FontTypeNames.FONTTYPE_TALK)
                            Exit Sub
    
                        End If
                    
140                     If ObjData(.Invent.Object(Slot).ObjIndex).Newbie = 1 Then
142                         Call WriteConsoleMsg(UserIndex, "No puedes comerciar objetos newbie.", e_FontTypeNames.FONTTYPE_TALK)
                            Exit Sub
                        End If
    
                    End If

                End If
            
                'Prevent offer changes (otherwise people would ripp off other players)
                'If .ComUsu.Objeto > 0 Then
                '     Call WriteConsoleMsg(UserIndex, "No podés cambiar tu oferta.", e_FontTypeNames.FONTTYPE_TALK)
                '     Exit Sub

                '  End If
            
                'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
144             If .flags.Navegando = 1 Then
146                 If .Invent.BarcoSlot = Slot Then
148                     Call WriteConsoleMsg(UserIndex, "No podés vender tu barco mientras lo estás usando.", e_FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
150             If .flags.Montado = 1 Then
152                 If .Invent.MonturaSlot = Slot Then
154                     Call WriteConsoleMsg(UserIndex, "No podés vender tu montura mientras la estás usando.", e_FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
156             .ComUsu.Objeto = Slot
158             .ComUsu.cant = amount
            
                'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
160             If UserList(tUser).ComUsu.Acepto Then
162                 UserList(tUser).ComUsu.Acepto = False
164                 Call WriteConsoleMsg(tUser, .Name & " ha cambiado su oferta.", e_FontTypeNames.FONTTYPE_TALK)

                End If
            
                Dim ObjAEnviar As t_Obj
                
166             ObjAEnviar.amount = amount

                'Si no es oro tmb le agrego el objInex
168             If Slot <> FLAGORO Then ObjAEnviar.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                'Llamos a la funcion
170             Call EnviarObjetoTransaccion(tUser, UserIndex, ObjAEnviar)

            End If

        End With

        Exit Sub

HandleUserCommerceOffer_Err:
172     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOffer", Erl)
174
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
        
104         otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptPeace", Erl)
116

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
        
104         otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectAlliance", Erl)
116

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
        
104         otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectPeace", Erl)
116

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
        
104         otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptAlliance", Erl)
116

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim proposal As String
            Dim errorStr As String
        
102         guild = Reader.ReadString8()
104         proposal = Reader.ReadString8()
        
106         If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, e_RELACIONES_GUILD.PAZ, proposal, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada", e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            End If

        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim proposal As String
            Dim errorStr As String
        
102         guild = Reader.ReadString8()
104         proposal = Reader.ReadString8()
        
106         If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, e_RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada", e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            End If

        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim errorStr As String
            Dim details  As String
        
102         guild = Reader.ReadString8()
        
104         details = modGuilds.r_VerPropuesta(UserIndex, guild, e_RELACIONES_GUILD.ALIADOS, errorStr)
        
106         If LenB(details) = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             Call WriteOfferDetails(UserIndex, details)

            End If

        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim errorStr As String
            Dim details  As String
        
102         guild = Reader.ReadString8()
        
104         details = modGuilds.r_VerPropuesta(UserIndex, guild, e_RELACIONES_GUILD.PAZ, errorStr)
        
106         If LenB(details) = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call WriteOfferDetails(UserIndex, details)

            End If
            
        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeaceDetails", Erl)
114

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim user    As String
            Dim details As String
        
102         user = Reader.ReadString8()
        
104         details = modGuilds.a_DetallesAspirante(UserIndex, user)
        
106         If LenB(details) = 0 Then
108             Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             Call WriteShowUserRequest(UserIndex, details)

            End If

        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestJoinerInfo", Erl)
114

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
    
100     Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, e_RELACIONES_GUILD.ALIADOS))
        
        Exit Sub

HandleGuildAlliancePropList_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAlliancePropList", Erl)
104
        
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

100     Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, e_RELACIONES_GUILD.PAZ))
        
        Exit Sub

HandleGuildPeacePropList_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
104
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild           As String
            Dim errorStr        As String
            Dim otherGuildIndex As Integer
        
102         guild = Reader.ReadString8()
        
104         otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
106         If otherGuildIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
                'WAR shall be!
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", e_FontTypeNames.FONTTYPE_GUILD))
114             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
116             Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

            End If

        End With
        
        Exit Sub
        
errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
120

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
        
        On Error GoTo errHandler

100     Call modGuilds.ActualizarWebSite(UserIndex, Reader.ReadString8())

        Exit Sub
        
errHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildNewWebsite", Erl)
104

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim errorStr As String
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
106             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
108             tUser = NameIndex(UserName)

110             If tUser > 0 Then
112                 Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
114                 Call RefreshCharStatus(tUser)

                End If
            
116             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("[" & UserName & "] ha sido aceptado como miembro del clan.", e_FontTypeNames.FONTTYPE_GUILD))
118             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

            End If

        End With
        
        Exit Sub
        
errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
122

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim errorStr As String
            Dim UserName As String
            Dim Reason   As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
106         If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             tUser = NameIndex(UserName)
            
112             If tUser > 0 Then
114                 Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, e_FontTypeNames.FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
116                 Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
120

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName   As String
            Dim GuildIndex As Integer
        
102         UserName = Reader.ReadString8()
        
104         GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
106         If GuildIndex > 0 Then
                Dim expulsadoIndex As Integer
108             expulsadoIndex = NameIndex(UserName)
110             If expulsadoIndex > 0 Then Call WriteConsoleMsg(expulsadoIndex, "Has sido expulsado del clan.", e_FontTypeNames.FONTTYPE_GUILD)
            
112             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", e_FontTypeNames.FONTTYPE_GUILD))
114             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Else
116             Call WriteConsoleMsg(UserIndex, "No podés expulsar ese personaje del clan.", e_FontTypeNames.FONTTYPE_GUILD)

            End If

        End With
        
        Exit Sub
        
errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildKickMember", Erl)
120

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
    
        On Error GoTo errHandler

100     Call modGuilds.ActualizarNoticias(UserIndex, Reader.ReadString8())

        Exit Sub
        
errHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildUpdateNews", Erl)
104

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
    
        On Error GoTo errHandler

100     Call modGuilds.SendDetallesPersonaje(UserIndex, Reader.ReadString8())

        Exit Sub
        
errHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberInfo", Erl)
104

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
100     With UserList(UserIndex)

            Dim Error As String
        
102         If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
104             Call WriteConsoleMsg(UserIndex, Error, e_FontTypeNames.FONTTYPE_GUILD)
            Else
106             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub

HandleGuildOpenElections_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOpenElections", Erl)
110
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild       As String
            Dim application As String
            Dim errorStr    As String
        
102         guild = Reader.ReadString8()
104         application = Reader.ReadString8()
        
106         If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", e_FontTypeNames.FONTTYPE_GUILD)

            End If

        End With
        
        Exit Sub
        
errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestMembership", Erl)
114

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

        On Error GoTo errHandler
 
100     Call modGuilds.SendGuildDetails(UserIndex, Reader.ReadString8())

        Exit Sub
        
errHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestDetails", Erl)
104

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
    
100     With UserList(UserIndex)

            Dim nombres As String
        
102         For i = 1 To LastUser

104             If UserList(i).flags.UserLogged Then
            
106                 If UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero) Then
108                     nombres = nombres & " - " & UserList(i).Name
                    End If

110                 Count = Count + 1

                End If

112         Next i
        
            'Get total time in seconds
114         Time = ((GetTickCount()) - tInicioServer) \ 1000
        
            'Get times in dd:hh:mm:ss format
116         UpTimeStr = (Time Mod 60) & " segundos."
118         Time = Time \ 60
        
120         UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
122         Time = Time \ 60
        
124         UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
126         Time = Time \ 24
        
128         If Time = 1 Then
130             UpTimeStr = Time & " día, " & UpTimeStr
            Else
132             UpTimeStr = Time & " días, " & UpTimeStr
    
            End If
    
134         Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, e_FontTypeNames.FONTTYPE_INFO)

136         If .flags.Privilegios And e_PlayerType.user Then
138             Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados.", e_FontTypeNames.FONTTYPE_INFOIAO)
140             Call WriteConsoleMsg(UserIndex, "Tiempo en línea: " & UpTimeStr & " Record de usuarios en simultaneo: " & RecordUsuarios & ".", e_FontTypeNames.FONTTYPE_INFOIAO)

            Else
142             Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados: " & nombres & ".", e_FontTypeNames.FONTTYPE_INFOIAO)
144             Call WriteConsoleMsg(UserIndex, "Tiempo en línea: " & UpTimeStr & " Record de usuarios en simultaneo: " & RecordUsuarios & ".", e_FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End With
        
        Exit Sub

HandleOnline_Err:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnline", Erl)
148
        
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
    
100     With UserList(UserIndex)

102         If .flags.Paralizado = 1 Then
104             Call WriteConsoleMsg(UserIndex, "No podés salir estando paralizado.", e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
            'exit secure commerce
106         If .ComUsu.DestUsu > 0 Then
108             tUser = .ComUsu.DestUsu
            
110             If UserList(tUser).flags.UserLogged Then
            
112                 If UserList(tUser).ComUsu.DestUsu = UserIndex Then
114                     Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", e_FontTypeNames.FONTTYPE_TALK)
116                     Call FinComerciarUsu(tUser)

                    End If

                End If
            
118             Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", e_FontTypeNames.FONTTYPE_TALK)
120             Call FinComerciarUsu(UserIndex)

        End If

138         Call Cerrar_Usuario(UserIndex)

        End With

        Exit Sub

HandleQuit_Err:
140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuit", Erl)
142
        
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
    
100     With UserList(UserIndex)

            'obtengo el guildindex
102         GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
104         If GuildIndex > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Dejas el clan.", e_FontTypeNames.FONTTYPE_GUILD)
108             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", e_FontTypeNames.FONTTYPE_GUILD))
            Else
110             Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de ningún clan.", e_FontTypeNames.FONTTYPE_GUILD)

            End If

        End With

        Exit Sub

HandleGuildLeave_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildLeave", Erl)
114
        
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
    
100     With UserList(UserIndex)

            'Dead people can't check their accounts
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
110         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 3 Then
112             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         Select Case NpcList(.flags.TargetNPC).NPCtype

                Case e_NPCType.Banquero
116                 Call WriteChatOverHead(UserIndex, "Tenes " & PonerPuntos(.Stats.Banco) & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
118             Case e_NPCType.Timbero

120                 If Not .flags.Privilegios And e_PlayerType.user Then
122                     earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
124                     If earnings >= 0 And Apuestas.Ganancias <> 0 Then
126                         percentage = Int(earnings * 100 / Apuestas.Ganancias)

                        End If
                    
128                     If earnings < 0 And Apuestas.Perdidas <> 0 Then
130                         percentage = Int(earnings * 100 / Apuestas.Perdidas)

                        End If
                    
132                     Call WriteConsoleMsg(UserIndex, "Entradas: " & PonerPuntos(Apuestas.Ganancias) & " Salida: " & PonerPuntos(Apuestas.Perdidas) & " Ganancia Neta: " & PonerPuntos(earnings) & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, e_FontTypeNames.FONTTYPE_INFO)

                    End If

            End Select

        End With
        
        Exit Sub

HandleRequestAccountState_Err:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAccountState", Erl)
136
        
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
        
100     With UserList(UserIndex)

            'Dead people can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
112             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's his pet
114         If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
            'Do it!
116         NpcList(.flags.TargetNPC).Movement = e_TipoAI.Estatico
        
118         Call Expresar(.flags.TargetNPC, UserIndex)

        End With
        
        Exit Sub

HandlePetStand_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetStand", Erl)
122
        
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
        
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
112             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make usre it's the user's pet
114         If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
            'Do it
116         Call FollowAmo(.flags.TargetNPC)
        
118         Call Expresar(.flags.TargetNPC, UserIndex)

        End With
        
        Exit Sub

HandlePetFollow_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetFollow", Erl)
122
        
End Sub

''
' Handles the "PetLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetLeave(ByVal UserIndex As Integer)
        '***************************************************
        
        On Error GoTo HandlePetLeave_Err
        
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make usre it's the user's pet
110         If NpcList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub

112         Call QuitarNPC(.flags.TargetNPC)

        End With
        
        Exit Sub

HandlePetLeave_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeave", Erl)
116
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If LenB(chat) <> 0 Then

                'Analize chat...
106             Call Statistics.ParseChat(chat)
            
108             If .Grupo.EnGrupo = True Then

                    Dim i As Byte
         
110                 For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
                    
                        'Call WriteConsoleMsg(UserList(.Grupo.Lider).Grupo.Miembros(i), "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
112                     Call WriteConsoleMsg(UserList(.Grupo.Lider).Grupo.Miembros(i), .Name & "> " & chat, e_FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
114                     Call WriteChatOverHead(UserList(.Grupo.Lider).Grupo.Miembros(i), chat, UserList(UserIndex).Char.CharIndex, &HFF8000)
                  
116                 Next i
            
                Else
                    'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_New_GRUPO)
118                 Call WriteConsoleMsg(UserIndex, "Grupo> No estas en ningun grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGrupoMsg", Erl)
122

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
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
112             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's the trainer
114         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Entrenador Then Exit Sub
        
116         Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)

        End With

        Exit Sub

HandleTrainList_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrainList", Erl)
120
        
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
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If HayOBJarea(.Pos, FOGATA) Then
108             Call WriteRestOK(UserIndex)
            
110             If Not .flags.Descansar Then
112                 Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzís a descansar.", e_FontTypeNames.FONTTYPE_INFO)
                Else
114                 Call WriteConsoleMsg(UserIndex, "Te levantas.", e_FontTypeNames.FONTTYPE_INFO)

                End If
            
116             .flags.Descansar = Not .flags.Descansar
            Else

118             If .flags.Descansar Then
120                 Call WriteRestOK(UserIndex)
122                 Call WriteConsoleMsg(UserIndex, "Te levantas.", e_FontTypeNames.FONTTYPE_INFO)
                
124                 .flags.Descansar = False
                    Exit Sub

                End If
            
126             Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        Exit Sub

HandleRest_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRest", Erl)
130
        
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
        
100     With UserList(UserIndex)

            'Si ya tiene el mana completo, no lo dejamos meditar.
102         If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
                           
            'Las clases NO MAGICAS no meditan...
104         If .clase = e_Class.Hunter Or .clase = e_Class.Trabajador Or .clase = e_Class.Warrior Or .clase = e_Class.Pirat Or .clase = e_Class.Thief Then Exit Sub

106         If .flags.Muerto = 1 Then
108             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
110         If .flags.Montado = 1 Then
112             Call WriteConsoleMsg(UserIndex, "No podes meditar estando montado.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

114         .flags.Meditando = Not .flags.Meditando

116         If .flags.Meditando Then

118             .Counters.TimerMeditar = 0

120             Select Case .Stats.ELV

                    Case 1 To 14
122                     .Char.FX = e_Meditaciones.MeditarInicial

124                 Case 15 To 29
126                     .Char.FX = e_Meditaciones.MeditarMayor15

128                 Case 30 To 39
130                     .Char.FX = e_Meditaciones.MeditarMayor30

132                 Case 40 To 44
134                     .Char.FX = e_Meditaciones.MeditarMayor40

136                 Case 45 To 46
138                     .Char.FX = e_Meditaciones.MeditarMayor45

140                 Case Else
142                     .Char.FX = e_Meditaciones.MeditarMayor47

                End Select

            Else
144             .Char.FX = 0

                'Call WriteLocaleMsg(UserIndex, "123", e_FontTypeNames.FONTTYPE_INFO)
            End If

146         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, .Char.FX))

        End With
        
        Exit Sub

HandleMeditate_Err:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMeditate", Erl)
150
        
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
100     With UserList(UserIndex)

            'Se asegura que el target es un npc
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
106         If (NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
            'Make sure it's close enough
108         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         Call RevivirUsuario(UserIndex)
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Curar, 100, False))
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
118         Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleResucitate_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResucitate", Erl)
122
        
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
100     With UserList(UserIndex)

            'Se asegura que el target es un npc
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If (NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Revividor And NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         .Stats.MinHp = .Stats.MaxHp
        
114         Call WriteUpdateHP(UserIndex)
        
116         Call WriteConsoleMsg(UserIndex, "ííHas sido curado!!", e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleHeal_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHeal", Erl)
120
        
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

100     Call SendUserStatsTxt(UserIndex, UserIndex)
        
        Exit Sub

HandleRequestStats_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestStats", Erl)
104
        
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

100     Call SendHelp(UserIndex)
        
        Exit Sub

HandleHelp_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHelp", Erl)
104
        
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
100     With UserList(UserIndex)

            'Dead people can't commerce
102         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Is it already in commerce mode??
106         If .flags.Comerciando Then
108             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
110         If .flags.TargetNPC > 0 Then
                
                'VOS, como GM, NO podes COMERCIAR con NPCs. (excepto Dioses y Admins)
112             If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
114                 Call WriteConsoleMsg(UserIndex, "No podés vender items.", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                'Does the NPC want to trade??
116             If NpcList(.flags.TargetNPC).Comercia = 0 Then
118                 If LenB(NpcList(.flags.TargetNPC).Desc) <> 0 Then
120                     Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If
                
                    Exit Sub

                End If
            
122             If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 3 Then
124                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Start commerce....
126             Call IniciarComercioNPC(UserIndex)
                
128         ElseIf .flags.TargetUser > 0 Then

                ' **********************  Comercio con Usuarios  *********************
                
                'VOS, como GM, NO podes COMERCIAR con usuarios. (excepto Dioses y Admins)
130             If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
132                 Call WriteConsoleMsg(UserIndex, "No podés vender items.", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                'NO podes COMERCIAR CON un GM. (excepto Dioses y Admins)
134             If (UserList(.flags.TargetUser).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
136                 Call WriteConsoleMsg(UserIndex, "No podés vender items a este usuario.", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                'Is the other one dead??
138             If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                    Call FinComerciarUsu(.flags.TargetUser, True)
140                 Call WriteConsoleMsg(UserIndex, "¡¡No podés comerciar con los muertos!!", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it me??
142             If .flags.TargetUser = UserIndex Then
144                 Call WriteConsoleMsg(UserIndex, "No podés comerciar con vos mismo...", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Check distance
146             If .Pos.Map <> UserList(.flags.TargetUser).Pos.Map Or Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                    Call FinComerciarUsu(.flags.TargetUser, True)
148                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del usuario.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
 
                'Check if map is not safe
                If MapInfo(.Pos.Map).Seguro = 0 Then
                    Call FinComerciarUsu(.flags.TargetUser, True)
                    Call WriteConsoleMsg(UserIndex, "No se puede usar el comercio seguro en zona insegura.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Is he already trading?? is it with me or someone else??
150             If UserList(.flags.TargetUser).flags.Comerciando = True Then
                    Call FinComerciarUsu(.flags.TargetUser, True)
152                 Call WriteConsoleMsg(UserIndex, "No podés comerciar con el usuario en este momento.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Initialize some variables...
154             .ComUsu.DestUsu = .flags.TargetUser
156             .ComUsu.DestNick = UserList(.flags.TargetUser).Name
158             .ComUsu.cant = 0
160             .ComUsu.Objeto = 0
162             .ComUsu.Acepto = False
            
                'Rutina para comerciar con otro usuario
164             Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)

            Else
166             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleCommerceStart_Err:
168     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceStart", Erl)
170
        
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
100     With UserList(UserIndex)

            'Dead people can't commerce
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If .flags.Comerciando Then
108             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
110         If .flags.TargetNPC > 0 Then
112             If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 6 Then
114                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'If it's the banker....
116             If NpcList(.flags.TargetNPC).NPCtype = e_NPCType.Banquero Then
118                 Call IniciarDeposito(UserIndex)

                End If

            Else
120             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleBankStart_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankStart", Erl)
124
        
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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteConsoleMsg(UserIndex, "Debes acercarte mís.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
116             Call EnlistarArmadaReal(UserIndex)
            
            Else
118             Call EnlistarCaos(UserIndex)

            End If

        End With
        
        Exit Sub

HandleEnlist_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEnlist", Erl)
122
        
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
100     With UserList(UserIndex)

            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
114             If .Faccion.ArmadaReal = 0 Then
116                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

118             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Else

120             If .Faccion.FuerzasCaos = 0 Then
122                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

124             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With
        
        Exit Sub

HandleInformation_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInformation", Erl)
128
        
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
100     With UserList(UserIndex)

            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
        
114             If .Faccion.ArmadaReal = 0 Then
116                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

118             Call RecompensaArmadaReal(UserIndex)
            
            Else

120             If .Faccion.FuerzasCaos = 0 Then
122                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

124             Call RecompensaCaos(UserIndex)

            End If

        End With
        
        Exit Sub

HandleReward_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReward", Erl)
128
        
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

100     Call SendMOTD(UserIndex)
        
        Exit Sub

HandleRequestMOTD_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMOTD", Erl)
104
        
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
100     Time = ((GetTickCount()) - tInicioServer) \ 1000
    
        'Get times in dd:hh:mm:ss format
102     UpTimeStr = (Time Mod 60) & " segundos."
104     Time = Time \ 60
    
106     UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
108     Time = Time \ 60
    
110     UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
112     Time = Time \ 24
    
114     If Time = 1 Then
116         UpTimeStr = Time & " día, " & UpTimeStr
        Else
118         UpTimeStr = Time & " días, " & UpTimeStr

        End If
    
120     Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, e_FontTypeNames.FONTTYPE_INFO)
        
        Exit Sub

HandleUpTime_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUpTime", Erl)
124
        
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

100     Call ConsultaPopular.SendInfoEncuesta(UserIndex)
        
        Exit Sub

HandleInquiry_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInquiry", Erl)
104
        
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
        
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If LenB(chat) <> 0 Then

                'Analize chat...
106             Call Statistics.ParseChat(chat)

                ' WyroX: Foto-denuncias - Push message
                Dim i As Integer

108             For i = 1 To UBound(.flags.ChatHistory) - 1
110                 .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                
112             .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
114             If .GuildIndex > 0 Then
                    'HarThaoS: si es leade mando un 10 para el status del color(medio villero pero me dio paja)
116                 If LCase(GuildLeader(.GuildIndex)) = .Name Then
118                     Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat, 10))
                    Else
120                     Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat, .Faccion.Status))
                    End If
                    'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                    'Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMessage", Erl)
124

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
100     Call CentinelaCheckClave(UserIndex, Reader.ReadInt16())

        Exit Sub

HandleCentinelReport_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCentinelReport", Erl)
104
        
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
100     With UserList(UserIndex)

            Dim onlineList As String
102             onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
104         If .GuildIndex <> 0 Then
106             Call WriteConsoleMsg(UserIndex, "Compañeros de tu clan conectados: " & onlineList, e_FontTypeNames.FONTTYPE_GUILDMSG)
            
            Else
108             Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", e_FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End With
        
        Exit Sub

HandleGuildOnline_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnline", Erl)
112
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If LenB(chat) <> 0 Then

                'Analize chat...
106             Call Statistics.ParseChat(chat)

                ' WyroX: Foto-denuncias - Push message
                Dim i As Long
108             For i = 1 To UBound(.flags.ChatHistory) - 1
110                 .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                
112             .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
114             If .flags.Privilegios And e_PlayerType.RoyalCouncil Then
116                 Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, e_FontTypeNames.FONTTYPE_CONSEJO))

118             ElseIf .flags.Privilegios And e_PlayerType.ChaosCouncil Then
120                 Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, e_FontTypeNames.FONTTYPE_CONSEJOCAOS))

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilMessage", Erl)
124

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim request As String
102             request = Reader.ReadString8()
        
104         If LenB(request) <> 0 Then
106             Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", e_FontTypeNames.FONTTYPE_INFO)
108             Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, e_FontTypeNames.FONTTYPE_GUILDMSG))

            End If

        End With
    
        Exit Sub
        
errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoleMasterRequest", Erl)
112

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Description As String
102             Description = Reader.ReadString8()
        
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "No podés cambiar la descripción estando muerto.", e_FontTypeNames.FONTTYPE_INFOIAO)

            Else
            
108             If Len(Description) > 128 Then
110                 Call WriteConsoleMsg(UserIndex, "La descripción es muy larga.", e_FontTypeNames.FONTTYPE_INFOIAO)

112             ElseIf Not DescripcionValida(Description) Then
114                 Call WriteConsoleMsg(UserIndex, "La descripción tiene carácteres inválidos.", e_FontTypeNames.FONTTYPE_INFOIAO)
                
                Else
116                 .Desc = Trim$(Description)
118                 Call WriteConsoleMsg(UserIndex, "La descripción a cambiado.", e_FontTypeNames.FONTTYPE_INFOIAO)

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeDescription", Erl)
122

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim vote     As String
            Dim errorStr As String
        
102         vote = Reader.ReadString8()
        
104         If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
106             Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
108             Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", e_FontTypeNames.FONTTYPE_GUILD)

            End If
 
        End With
        
        Exit Sub
        
errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildVote", Erl)
112

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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Name As String
102             Name = Reader.ReadString8()

            ' Si un GM usa este comando, me fijo que me haya dado el nick del PJ a analizar.
104         If EsGM(UserIndex) And LenB(Name) = 0 Then Exit Sub
        
            Dim Count As Integer

106         If (InStrB(Name, "\") <> 0) Then
108             Name = Replace(Name, "\", vbNullString)

            End If

110         If (InStrB(Name, "/") <> 0) Then
112             Name = Replace(Name, "/", vbNullString)

            End If

114         If (InStrB(Name, ":") <> 0) Then
116             Name = Replace(Name, ":", vbNullString)

            End If

118         If (InStrB(Name, "|") <> 0) Then
120             Name = Replace(Name, "|", vbNullString)

            End If
           
            Dim TargetUserName As String

122         If EsGM(UserIndex) Then
        
124             If PersonajeExiste(Name) Then
126                 TargetUserName = Name
                
                Else
128                 Call WriteConsoleMsg(UserIndex, "El personaje " & TargetUserName & " no existe.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
            Else
        
130             TargetUserName = .Name
            
            End If

134         Count = GetUserAmountOfPunishmentsDatabase(TargetUserName)


138         If Count = 0 Then
140             Call WriteConsoleMsg(UserIndex, "Sin prontuario..", e_FontTypeNames.FONTTYPE_INFO)
            Else
144             Call SendUserPunishmentsDatabase(UserIndex, TargetUserName)
            End If

        End With
    
        Exit Sub
    
errHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePunishments", Erl)
154

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

100     With UserList(UserIndex)

            Dim amount As Integer
102             amount = Reader.ReadInt16()
        
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                
108         ElseIf .flags.TargetNPC = 0 Then
                'Validate target NPC
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)

112         ElseIf Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
114             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                
116         ElseIf NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Timbero Then
118             Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

120         ElseIf amount < 1 Then
122             Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

124         ElseIf amount > 10000 Then
126             Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 10.000 monedas.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

128         ElseIf .Stats.GLD < amount Then
130             Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            Else

132             If RandomNumber(1, 100) <= 45 Then
134                 .Stats.GLD = .Stats.GLD + amount
136                 Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & PonerPuntos(amount) & " monedas de oro!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
138                 Apuestas.Perdidas = Apuestas.Perdidas + amount
140                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
142                 .Stats.GLD = .Stats.GLD - amount
144                 Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & PonerPuntos(amount) & " monedas de oro.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
146                 Apuestas.Ganancias = Apuestas.Ganancias + amount
148                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If
            
150             Apuestas.Jugadas = Apuestas.Jugadas + 1
            
152             Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
154             Call WriteUpdateGold(UserIndex)

            End If

        End With

        Exit Sub

HandleGamble_Err:
156     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGamble", Erl)
158
        
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

100     With UserList(UserIndex)

            Dim opt As Byte
102             opt = Reader.ReadInt8()
        
104         Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), e_FontTypeNames.FONTTYPE_GUILD)

        End With
        
        Exit Sub

HandleInquiryVote_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInquiryVote", Erl)
108
        
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

100     With UserList(UserIndex)

            Dim amount As Long
102             amount = Reader.ReadInt32()
        
            'Dead people can't leave a faction.. they can't talk...
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then Exit Sub
        
114         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If amount > 0 And amount <= .Stats.Banco Then
120             .Stats.Banco = .Stats.Banco - amount
122             .Stats.GLD = .Stats.GLD + amount
                'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

124             Call WriteUpdateGold(UserIndex)
126             Call WriteGoliathInit(UserIndex)

            Else
128             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With

        Exit Sub

HandleBankExtractGold_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractGold", Erl)
132
        
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
        
100     With UserList(UserIndex)

            'Dead people can't leave a faction.. they can't talk...
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If .Faccion.ArmadaReal = 0 And .Faccion.FuerzasCaos = 0 Then
108             If .Faccion.Status = 1 Then
110                 Call VolverCriminal(UserIndex)
112                 Call WriteConsoleMsg(UserIndex, "Ahora sos un criminal.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             If .Faccion.ArmadaReal = 1 Then
118                 Call WriteConsoleMsg(UserIndex, "Para salir del ejercito debes ir a visitar al rey.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
120             ElseIf .Faccion.FuerzasCaos = 1 Then
122                 Call WriteConsoleMsg(UserIndex, "Para salir de la legion debes ir a visitar al diablo.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Exit Sub
            End If
        
124         If NpcList(.flags.TargetNPC).NPCtype = e_NPCType.Enlistador Then
                'Quit the Royal Army?
126             If .Faccion.ArmadaReal = 1 Then
128                 If NpcList(.flags.TargetNPC).flags.Faccion = 0 Then
                
                        'HarThaoS
                        'Si tiene clan
130                     If .GuildIndex > 0 Then
                            'Y no es leader
132                         If Not PersonajeEsLeader(.Name) Then
                                'Lo echo a la verga
134                             Call m_EcharMiembroDeClan(UserIndex, .Name)
136                             Call WriteConsoleMsg(UserIndex, "Has dejado el clan.", e_FontTypeNames.FONTTYPE_GUILD)
                            Else
138                             Call WriteChatOverHead(UserIndex, "Para dejar la facción primero deberás ceder el clan", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Exit Sub
                            End If
                        End If
                    
140                     Call ExpulsarFaccionReal(UserIndex)
142                     Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                        Exit Sub
                    Else
144                     Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                   
                    End If

                    'Quit the Chaos Legion??
146             ElseIf .Faccion.FuerzasCaos = 1 Then

148                 If NpcList(.flags.TargetNPC).flags.Faccion = 2 Then
                        'HarThaoS
                        'Si tiene clan
150                     If .GuildIndex > 0 Then
                            'Y no es leader
152                         If Not PersonajeEsLeader(.Name) Then
                                'Lo echo a la verga
154                             Call m_EcharMiembroDeClan(UserIndex, .Name)
156                             Call WriteConsoleMsg(UserIndex, "Has dejado el clan.", e_FontTypeNames.FONTTYPE_GUILD)
                            Else
158                             Call WriteChatOverHead(UserIndex, "Para dejar la facción primero deberás ceder el clan", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Exit Sub
                            End If
                        End If
                    
160                     Call ExpulsarFaccionCaos(UserIndex)
162                     Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
164                     Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If

                Else
166                 Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

                End If

            End If
    
        End With
        
        Exit Sub

HandleLeaveFaction_Err:
168     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeaveFaction", Erl)
170
        
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

100     With UserList(UserIndex)

            Dim amount As Long
102             amount = Reader.ReadInt32()
        
            'Dead people can't leave a faction.. they can't talk...
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
112         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then Exit Sub
        
114         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If amount > 0 And amount <= .Stats.GLD Then
120             .Stats.Banco = .Stats.Banco + amount
122             .Stats.GLD = .Stats.GLD - amount
                'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
124             Call WriteUpdateGold(UserIndex)
126             Call WriteGoliathInit(UserIndex)
            Else
128             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With
        
        Exit Sub

HandleBankDepositGold_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDepositGold", Erl)
132
        
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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

106         If EventoActivo Then
108             Call FinalizarEvento
            Else
110             Call WriteConsoleMsg(UserIndex, "No hay ningun evento activo.", e_FontTypeNames.FONTTYPE_New_Eventos)
        
            End If
        
        End With
        
        Exit Sub

HandleDenounce_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
114
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild       As String
            Dim memberCount As Integer
            Dim i           As Long
            Dim UserName    As String
        
102         guild = Reader.ReadString8()
        
104         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios) Then

106             If (InStrB(guild, "\") <> 0) Then
108                 guild = Replace(guild, "\", "")

                End If

110             If (InStrB(guild, "/") <> 0) Then
112                 guild = Replace(guild, "/", "")

                End If
            
114             If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
116                 Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, e_FontTypeNames.FONTTYPE_INFO)

                Else
                
118                 memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
120                 For i = 1 To memberCount
122                     UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
124                     Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", e_FontTypeNames.FONTTYPE_INFO)
126                 Next i

                End If
        
            End If
            
        End With
        
        Exit Sub
        
errHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberList", Erl)
130

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

        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim Message As String
102             Message = Reader.ReadString8()

104         If EsGM(UserIndex) Then
106             Call LogGM(.Name, "Mensaje a Gms: " & Message)
        
108             If LenB(Message) <> 0 Then
                    'Analize chat...
110                 Call Statistics.ParseChat(Message)
            
112                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " » " & message, e_FontTypeNames.FONTTYPE_GMMSG))

                End If

            End If

        End With

        Exit Sub
    
errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMMessage", Erl)
116

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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
        
104             .showName = Not .showName 'Show / Hide the name
            
106             Call RefreshCharStatus(UserIndex)

            End If

        End With
        
        Exit Sub

HandleShowName_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowName", Erl)
110
        
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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
            Dim i    As Long
            Dim list As String

104         For i = 1 To LastUser

106             If UserList(i).ConnIDValida Then
108                 If UserList(i).Faccion.ArmadaReal = 1 Then
110                     If UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                         list = list & UserList(i).Name & ", "

                        End If

                    End If

                End If

114         Next i

        End With
    
116     If Len(list) > 0 Then
118         Call WriteConsoleMsg(UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), e_FontTypeNames.FONTTYPE_INFO)
        Else
120         Call WriteConsoleMsg(UserIndex, "No hay Armadas conectados", e_FontTypeNames.FONTTYPE_INFO)

        End If
        
        Exit Sub

HandleOnlineRoyalArmy_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineRoyalArmy", Erl)
124
        
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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
            Dim i    As Long
            Dim list As String

104         For i = 1 To LastUser

106             If UserList(i).ConnIDValida Then
108                 If UserList(i).Faccion.FuerzasCaos = 1 Then
110                     If UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                         list = list & UserList(i).Name & ", "

                        End If

                    End If

                End If

114         Next i

        End With

116     If Len(list) > 0 Then
118         Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), e_FontTypeNames.FONTTYPE_INFO)
        
        Else
120         Call WriteConsoleMsg(UserIndex, "No hay Caos conectados", e_FontTypeNames.FONTTYPE_INFO)

        End If
        
        Exit Sub

HandleOnlineChaosLegion_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineChaosLegion", Erl)
124
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
102             UserName = Reader.ReadString8()
        
            Dim tIndex As Integer

            Dim x      As Long
            Dim Y      As Long

            Dim i      As Long
            
            Dim Found  As Boolean
        
104         If Not EsGM(UserIndex) Then Exit Sub
        
            'Check the user has enough powers
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or Ayuda.Existe(UserName) Then
108             tIndex = NameIndex(UserName)
            
110             If tIndex <= 0 Then
                    ' Si está offline, comparamos privilegios offline, para no revelar si está el gm conectado
112                 If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(UserName)) >= 0 Then
114                     Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
116                     Call WriteConsoleMsg(UserIndex, "No podés ir cerca de un GM de mayor jerarquía.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
118                 If CompararPrivilegiosUser(UserIndex, tIndex) >= 0 Then
120                     For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
122                         For x = UserList(tIndex).Pos.x - i To UserList(tIndex).Pos.x + i
124                             For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

126                                 If MapData(UserList(tIndex).Pos.Map, x, Y).UserIndex = 0 Then
128                                     If LegalPos(UserList(tIndex).Pos.Map, x, Y, True, True) Then
130                                         Call WriteConsoleMsg(UserIndex, "Te teletransportaste cerca de " & UserList(tIndex).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
132                                         Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, x, Y, True)
134                                         Found = True
                                            Exit For
                                        End If

                                    End If

136                             Next Y
                            
138                             If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
140                         Next x
                        
142                         If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
144                     Next i
                    
                        'No space found??
146                     If Not Found Then
148                         Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
150                     Call WriteConsoleMsg(UserIndex, "No podés ir cerca de un GM de mayor jerarquía.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
152             Call WriteConsoleMsg(UserIndex, "Servidor » No podés ir cerca de ningun Usuario si no pidio SOS.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub
        
errHandler:
154     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoNearby", Erl)
156

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim comment As String
102             comment = Reader.ReadString8()
        
104         If Not .flags.Privilegios And e_PlayerType.user Then
106             Call LogGM(.Name, "Comentario: " & comment)
108             Call WriteConsoleMsg(UserIndex, "Comentario salvado...", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub
        
errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComment", Erl)
112

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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
104         Call LogGM(.Name, "Hora.")

        End With
    
106     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, e_FontTypeNames.FONTTYPE_INFO))
        
        Exit Sub

HandleServerTime_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerTime", Erl)
110
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.user)) = 0 Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                Else

112                 If CompararPrivilegiosUser(UserIndex, tUser) >= 0 Then
114                     Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.x & ", " & UserList(tUser).Pos.Y & ".", e_FontTypeNames.FONTTYPE_INFO)
116                     Call LogGM(.Name, "/Donde " & UserName)

                    End If

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub
        
errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhere", Erl)
122

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
        'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
        '***************************************************

100     With UserList(UserIndex)

            Dim Map As Integer
            Dim i, j As Long
            Dim NPCcount1, NPCcount2 As Integer
            Dim NPCcant1() As Integer
            Dim NPCcant2() As Integer
            Dim List1()    As String
            Dim List2()    As String
        
102         Map = Reader.ReadInt16()
        
104         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
106         If MapaValido(Map) Then

108             For i = 1 To LastNPC

                    'VB isn't lazzy, so we put more restrictive condition first to speed up the process
110                 If NpcList(i).Pos.Map = Map Then

                        'íesta vivo?
112                     If NpcList(i).flags.NPCActive And NpcList(i).Hostile = 1 Then
114                         If NPCcount1 = 0 Then
116                             ReDim List1(0) As String
118                             ReDim NPCcant1(0) As Integer
120                             NPCcount1 = 1
122                             List1(0) = NpcList(i).Name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
124                             NPCcant1(0) = 1
                            Else

126                             For j = 0 To NPCcount1 - 1

128                                 If Left$(List1(j), Len(NpcList(i).Name)) = NpcList(i).Name Then
130                                     List1(j) = List1(j) & ", (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
132                                     NPCcant1(j) = NPCcant1(j) + 1
                                        Exit For

                                    End If

134                             Next j

136                             If j = NPCcount1 Then
138                                 ReDim Preserve List1(0 To NPCcount1) As String
140                                 ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
142                                 NPCcount1 = NPCcount1 + 1
144                                 List1(j) = NpcList(i).Name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
146                                 NPCcant1(j) = 1

                                End If

                            End If

                        Else

148                         If NPCcount2 = 0 Then
150                             ReDim List2(0) As String
152                             ReDim NPCcant2(0) As Integer
154                             NPCcount2 = 1
156                             List2(0) = NpcList(i).Name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
158                             NPCcant2(0) = 1
                            Else

160                             For j = 0 To NPCcount2 - 1

162                                 If Left$(List2(j), Len(NpcList(i).Name)) = NpcList(i).Name Then
164                                     List2(j) = List2(j) & ", (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
166                                     NPCcant2(j) = NPCcant2(j) + 1
                                        Exit For

                                    End If

168                             Next j

170                             If j = NPCcount2 Then
172                                 ReDim Preserve List2(0 To NPCcount2) As String
174                                 ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
176                                 NPCcount2 = NPCcount2 + 1
178                                 List2(j) = NpcList(i).Name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.Y & ")"
180                                 NPCcant2(j) = 1

                                End If

                            End If

                        End If

                    End If

182             Next i
            
184             Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", e_FontTypeNames.FONTTYPE_WARNING)

186             If NPCcount1 = 0 Then
188                 Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", e_FontTypeNames.FONTTYPE_INFO)
                Else

190                 For j = 0 To NPCcount1 - 1
192                     Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), e_FontTypeNames.FONTTYPE_INFO)
194                 Next j

                End If

196             Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", e_FontTypeNames.FONTTYPE_WARNING)

198             If NPCcount2 = 0 Then
200                 Call WriteConsoleMsg(UserIndex, "No hay mís NPCS", e_FontTypeNames.FONTTYPE_INFO)
                Else

202                 For j = 0 To NPCcount2 - 1
204                     Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), e_FontTypeNames.FONTTYPE_INFO)
206                 Next j

                End If

208             Call LogGM(.Name, "Numero enemigos en mapa " & Map)

            End If

        End With
        
        Exit Sub

HandleCreaturesInMap_Err:
210     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreaturesInMap", Erl)
212
        
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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
104         Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
        
106         Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

        End With
        
        Exit Sub

HandleWarpMeToTarget_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpMeToTarget", Erl)
110
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Map      As Integer
            Dim x        As Byte
            Dim Y        As Byte
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Map = Reader.ReadInt16()
106         x = Reader.ReadInt8()
108         Y = Reader.ReadInt8()

110         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
            
112         If .flags.Privilegios And e_PlayerType.Consejero Then
        
114             If MapInfo(Map).Seguro = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ciudades.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    'Si manda yo o su propio nombre
118             ElseIf LCase$(UserName) <> LCase$(UserList(UserIndex).Name) And UCase$(UserName) <> "YO" Then
120                 Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ti mismo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
            
            '¿Para que te vas a transportar a la misma posicion?
122         If .Pos.Map = Map And .Pos.x = x And .Pos.Y = Y Then Exit Sub
            
124         If MapaValido(Map) And LenB(UserName) <> 0 Then

126             If UCase$(UserName) <> "YO" Then
128                 tUser = NameIndex(UserName)
                
                Else
130                 tUser = UserIndex

                End If
            
132             If tUser <= 0 Then
134                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)

136             ElseIf InMapBounds(Map, x, Y) Then
138                 Call FindLegalPos(tUser, Map, x, Y)
140                 Call WarpUserChar(tUser, Map, x, Y, True)

142                 If tUser <> UserIndex Then
144                     Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & x & " Y:" & Y)
                    End If
                        
                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpChar", Erl)
148

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim minutos  As Integer
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         minutos = Reader.ReadInt16()

106         If EsGM(UserIndex) Then
108             tUser = NameIndex(UserName)
        
110             If tUser <= 0 Then

112                 If PersonajeExiste(UserName) Then

114                     If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(UserName)) > 0 Then

116                         If minutos > 0 Then
118                             Call SilenciarUserDatabase(UserName, minutos)
120                             Call SavePenaDatabase(UserName, .Name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
122                             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha silenciado a " & UserName & "(offline) por " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_GM))
124                             Call LogGM(.Name, "Silenciar a " & UserList(tUser).Name & " por " & minutos & " minutos.")
                            Else
126                             Call DesilenciarUserDatabase(UserName)
128                             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha desilenciado a " & UserName & "(offline).", e_FontTypeNames.FONTTYPE_GM))
130                             Call LogGM(.Name, "Desilenciar a " & UserList(tUser).Name & ".")

                            End If
                            
                        Else
                        
132                         Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                    
134                     Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                
136             ElseIf CompararPrivilegiosUser(UserIndex, tUser) > 0 Then

138                 If minutos > 0 Then
140                     UserList(tUser).flags.Silenciado = 1
142                     UserList(tUser).flags.MinutosRestantes = minutos
144                     UserList(tUser).flags.SegundosPasados = 0

146                     Call SavePenaDatabase(UserName, .Name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
148                     Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha silenciado a " & UserList(tUser).Name & " por " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_GM))
150                     Call WriteConsoleMsg(tUser, "Has sido silenciado por los administradores, no podrás hablar con otros usuarios. Utilice /GM para pedir ayuda.", e_FontTypeNames.FONTTYPE_GM)
152                     Call LogGM(.Name, "Silenciar a " & UserList(tUser).Name & " por " & minutos & " minutos.")

                    Else
                    
154                     UserList(tUser).flags.Silenciado = 1

156                     Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha desilenciado a " & UserList(tUser).Name & ".", e_FontTypeNames.FONTTYPE_GM))
158                     Call WriteConsoleMsg(tUser, "Has sido desilenciado.", e_FontTypeNames.FONTTYPE_GM)
160                     Call LogGM(.Name, "Desilenciar a " & UserList(tUser).Name & ".")

                    End If
                    
                Else
                
162                 Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With
        
        Exit Sub
        
errHandler:
164     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSilence", Erl)
166

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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub

104         Call WriteShowSOSForm(UserIndex)

        End With
        
        Exit Sub

HandleSOSShowList_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSShowList", Erl)
108
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
102             UserName = Reader.ReadString8()
        
104         If Not .flags.Privilegios And e_PlayerType.user Then Call Ayuda.Quitar(UserName)

        End With
        
        Exit Sub
        
errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSRemove", Erl)
108

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
            Dim x        As Byte
            Dim Y        As Byte
        
102         UserName = Reader.ReadString8()

104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If LenB(UserName) <> 0 Then
108                 tUser = NameIndex(UserName)
                    
110                 If tUser <= 0 Then
112                     Call WriteConsoleMsg(UserIndex, "El jugador no está online.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                Else
114                 tUser = .flags.TargetUser

116                 If tUser <= 0 Then Exit Sub

                End If
      
118             If CompararPrivilegiosUser(tUser, UserIndex) > 0 Then
120                 Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser).Name & " que quieres ir a su posición.", e_FontTypeNames.FONTTYPE_INFO)
122                 Call WriteConsoleMsg(tUser, .Name & " quiere transportarse a tu ubicación. Escribe /sum " & .Name & " para traerlo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

124             x = UserList(tUser).Pos.x
126             Y = UserList(tUser).Pos.Y + 1

128             Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, x, Y)
                
130             Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, x, Y, True)
                    
132             If .flags.AdminInvisible = 0 Then
134                 Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", e_FontTypeNames.FONTTYPE_INFO)

                End If
                
136             Call WriteConsoleMsg(UserIndex, "Te has transportado hacia " & UserList(tUser).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
                    
138             Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.x & " Y:" & UserList(tUser).Pos.Y)
            Else
140             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo. solo puedes ir a Usuarios que piden SOS.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub
        
errHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoToChar", Erl)
144

End Sub

Private Sub HandleDarLlaveAUsuario(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String, tUser As Integer, Llave As Integer
        
102         UserName = Reader.ReadString8()
104         Llave = Reader.ReadInt16()
        
            ' Solo dios o admin
106         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then

                ' Me aseguro que esté activada la db
108             If Not Database_Enabled Then
110                 Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", e_FontTypeNames.FONTTYPE_INFO)
            
                    ' Me aseguro que el objeto sea una llave válida
112             ElseIf Llave < 1 Or Llave > NumObjDatas Then
114                 Call WriteConsoleMsg(UserIndex, "El número ingresado no es el de una llave válida.", e_FontTypeNames.FONTTYPE_INFO)

116             ElseIf ObjData(Llave).OBJType <> e_OBJType.otLlaves Then ' vb6 no tiene short-circuit evaluation :(
118                 Call WriteConsoleMsg(UserIndex, "El número ingresado no es el de una llave válida.", e_FontTypeNames.FONTTYPE_INFO)

                Else
                
120                 tUser = NameIndex(UserName)
                
122                 If tUser > 0 Then

                        ' Es un user online, guardamos la llave en la db
124                     If DarLlaveAUsuarioDatabase(UserName, Llave) Then

                            ' Actualizamos su llavero
126                         If MeterLlaveEnLLavero(tUser, Llave) Then
128                             Call WriteConsoleMsg(UserIndex, "Llave número " & Llave & " entregada a " & UserList(tUser).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
                            Else
130                             Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. El usuario no tiene más espacio en su llavero.", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        Else
132                         Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible.", e_FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                    Else
                    
                        ' No es un usuario online, nos fijamos si es un email
134                     If CheckMailString(UserName) Then

                            ' Es un email, intentamos guardarlo en la db
136                         If DarLlaveACuentaDatabase(UserName, Llave) Then
138                             Call WriteConsoleMsg(UserIndex, "Llave número " & Llave & " entregada a " & LCase$(UserName) & ".", e_FontTypeNames.FONTTYPE_INFO)
                            Else
140                             Call WriteConsoleMsg(UserIndex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible y que el email sea correcto.", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        Else
142                         Call WriteConsoleMsg(UserIndex, "El usuario no está online. Ingrese el email de la cuenta para otorgar la llave offline.", e_FontTypeNames.FONTTYPE_INFO)

                        End If
    
                    End If
                
144                 Call LogGM(.Name, "/DARLLAVE " & UserName & " " & Llave)

                End If

            Else
146             Call WriteConsoleMsg(UserIndex, "Servidor » Solo Dios y Administrador pueden dar llaves.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub
        
errHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDarLlaveAUsuario", Erl)
150

End Sub

Private Sub HandleSacarLlave(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSacarLlave_Err

100     With UserList(UserIndex)

            Dim Llave As Integer
102             Llave = Reader.ReadInt16()
        
            ' Solo dios o admin
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) Then

                ' Me aseguro que esté activada la db
106             If Not Database_Enabled Then
108                 Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", e_FontTypeNames.FONTTYPE_INFO)

                Else

                    ' Intento borrarla de la db
110                 If SacarLlaveDatabase(Llave) Then
112                     Call WriteConsoleMsg(UserIndex, "La llave " & Llave & " fue removida.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
114                     Call WriteConsoleMsg(UserIndex, "No se pudo sacar la llave. Asegúrese de que esté en uso.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

116                 Call LogGM(.Name, "/SACARLLAVE " & Llave)

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Solo Dios y Administrador pueden Sacar llaves.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub

HandleSacarLlave_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSacarLlave", Erl)
122
        
End Sub

Private Sub HandleVerLlaves(ByVal UserIndex As Integer)
        
        On Error GoTo HandleVerLlaves_Err

100     With UserList(UserIndex)

            ' Sólo GMs
102         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                ' Me aseguro que esté activada la db
104             If Not Database_Enabled Then
106                 Call WriteConsoleMsg(UserIndex, "Es necesario que el juego esté corriendo con base de datos.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                ' Leo y muestro todas las llaves usadas
108             Call VerLlavesDatabase(UserIndex)
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor » Solo Dios y Administrador pueden ver llaves.", e_FontTypeNames.FONTTYPE_INFO)
            End If
                
        End With

        Exit Sub

HandleVerLlaves_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleVerLlaves", Erl)
114
        
End Sub

Private Sub HandleUseKey(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUseKey_Err

100     With UserList(UserIndex)

            Dim Slot As Byte
102             Slot = Reader.ReadInt8

104         Call UsarLlave(UserIndex, Slot)
                
        End With
        
        Exit Sub

HandleUseKey_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseKey", Erl)
108
        
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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
104         Call DoAdminInvisible(UserIndex)

        End With
        
        Exit Sub

HandleInvisible_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvisible", Erl)
108
        
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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
104         Call WriteShowGMPanelForm(UserIndex)

        End With
        
        Exit Sub

HandleGMPanel_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMPanel", Erl)
108
        
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
    
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         ReDim names(1 To LastUser) As String
108         Count = 1
        
110         For i = 1 To LastUser

112             If (LenB(UserList(i).Name) <> 0) Then
                
114                 names(Count) = UserList(i).Name
116                 Count = Count + 1
 
                End If

118         Next i
        
120         If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

        End With
        
        Exit Sub

HandleRequestUserList_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestUserList", Erl)
124
        
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
    
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » /TRABAJANDO es un comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         For i = 1 To LastUser

108             If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
110                 Users = Users & ", " & UserList(i).Name
                
                    ' Display the user being checked by the centinel
112                 If modCentinela.Centinela.RevisandoUserIndex = i Then Users = Users & " (*)"

                End If

114         Next i
        
116         If LenB(Users) <> 0 Then
118             Users = Right$(Users, Len(Users) - 2)
120             Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, e_FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleWorking_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorking", Erl)
126
        
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
    
100     With UserList(UserIndex)

        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         For i = 1 To LastUser

108             If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
110                 Users = Users & UserList(i).Name & ", "

                End If

112         Next i
        
114         If LenB(Users) <> 0 Then
116             Users = Left$(Users, Len(Users) - 2)
118             Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, e_FontTypeNames.FONTTYPE_INFO)
            Else
120             Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleHiding_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
124
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

        
        
        
            Dim UserName As String
            Dim Reason   As String
            Dim jailTime As Byte
            Dim Count    As Byte
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         jailTime = Reader.ReadInt8()
        
108         If InStr(1, UserName, "+") Then
110             UserName = Replace(UserName, "+", " ")

            End If
        
            '/carcel nick@motivo@<tiempo>
112         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then

114             If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", e_FontTypeNames.FONTTYPE_INFO)
                Else
118                 tUser = NameIndex(UserName)
                
120                 If tUser <= 0 Then
122                     Call WriteConsoleMsg(UserIndex, "El usuario no está online.", e_FontTypeNames.FONTTYPE_INFO)
                    Else

124                     If EsGM(tUser) Then
126                         Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", e_FontTypeNames.FONTTYPE_INFO)
                    
128                     ElseIf jailTime > 60 Then
130                         Call WriteConsoleMsg(UserIndex, "No podés encarcelar por más de 60 minutos.", e_FontTypeNames.FONTTYPE_INFO)
                        Else

132                         If (InStrB(UserName, "\") <> 0) Then
134                             UserName = Replace(UserName, "\", "")

                            End If

136                         If (InStrB(UserName, "/") <> 0) Then
138                             UserName = Replace(UserName, "/", "")

                            End If
                        
140                         If PersonajeExiste(UserName) Then
144                                 Call SavePenaDatabase(UserName, .Name & ": CARCEL " & jailTime & "m, MOTIVO: " & Reason & " " & Date & " " & Time)
                            End If
                        
152                         Call Encarcelar(tUser, jailTime, .Name)
154                         Call LogGM(.Name, " encarceló a " & UserName)

                        End If

                    End If

                End If
            Else
156             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub
        
errHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
160

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
100     With UserList(UserIndex)

102         If .flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero) Then
104             Call WriteConsoleMsg(UserIndex, "Solo los Administradores y Dioses pueden usar este comando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
        
            'Si estamos en el mapa pretoriano...
106         If .Pos.Map = MAPA_PRETORIANO Then

                '... solo los Dioses y Administradores pueden usar este comando en el mapa pretoriano.
108             If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios) = 0 Then
                
110                 Call WriteConsoleMsg(UserIndex, "Solo los Administradores y Dioses pueden usar este comando en el mapa pretoriano.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
            Dim tNPC As Integer
112         tNPC = .flags.TargetNPC
        
114         If tNPC > 0 Then

116             Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & NpcList(tNPC).Name, e_FontTypeNames.FONTTYPE_INFO)
            
                Dim auxNPC As t_Npc
118             auxNPC = NpcList(tNPC)
            
120             Call QuitarNPC(tNPC)
122             Call ReSpawnNpc(auxNPC)
            
            Else
124             Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleKillNPC_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPC", Erl)

128
        
End Sub

''
' Handles the "WarnUser" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Reason   As String

102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
            ' Tenes que ser Admin, Dios o Semi-Dios
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            ' Me fijo que esten todos los parametros.
110         If Len(UserName) = 0 Or Len(Trim$(Reason)) = 0 Then
112             Call WriteConsoleMsg(UserIndex, "Formato inválido. /advertencia nick@motivo", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Dim tUser As Integer
114         tUser = NameIndex(UserName)
        
            ' No advertir a GM's
116         If EsGM(tUser) Then
118             Call WriteConsoleMsg(UserIndex, "No podes advertir a Game Masters.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
120         If (InStrB(UserName, "\") <> 0) Then
122             UserName = Replace(UserName, "\", "")

            End If

124         If (InStrB(UserName, "/") <> 0) Then
126             UserName = Replace(UserName, "/", "")

            End If
                    
128         If PersonajeExiste(UserName) Then

132             Call SaveWarnDatabase(UserName, "ADVERTENCIA: " & Reason & " " & Date & " " & Time, .Name)

            
                ' Para el GM
140             Call WriteConsoleMsg(UserIndex, "Has advertido a " & UserName, e_FontTypeNames.FONTTYPE_CENTINELA)
142             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " ha advertido a " & UserName & " por " & Reason, e_FontTypeNames.FONTTYPE_GM))
144             Call LogGM(.Name, " advirtio a " & UserName & " por " & Reason)

                ' Si esta online...
146             If tUser >= 0 Then
                    ' Actualizo el valor en la memoria.
148                 UserList(tUser).Stats.Advertencias = UserList(tUser).Stats.Advertencias + 1
                
                    ' Para el usuario advertido
150                 Call WriteConsoleMsg(tUser, "Has sido advertido por " & Reason, e_FontTypeNames.FONTTYPE_CENTINELA)
152                 Call WriteConsoleMsg(tUser, "Tenés " & UserList(tUser).Stats.Advertencias & " advertencias actualmente.", e_FontTypeNames.FONTTYPE_CENTINELA)
                
                    ' Cuando acumulas cierta cantidad de advertencias...
154                 Select Case UserList(tUser).Stats.Advertencias
                
                        Case 3
156                         Call Encarcelar(tUser, 30, "Servidor")
                    
158                     Case 5
                            ' TODO: Banear PJ alv.
                    
                    End Select
                
                End If

            End If
        
        End With
    
        Exit Sub
    
errHandler:

160     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarnUser", Erl)

162

End Sub

Private Sub HandleMensajeUser(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Mensaje  As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Mensaje = Reader.ReadString8()
        
106         If EsGM(UserIndex) Then
        
108             If LenB(UserName) = 0 Or LenB(Mensaje) = 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Utilice /MENSAJEINFORMACION nick@mensaje", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
112                 tUser = NameIndex(UserName)
                
114                 If tUser Then
116                     Call WriteConsoleMsg(tUser, "Mensaje recibido de " & .Name & " [Game Master]:", e_FontTypeNames.FONTTYPE_New_DONADOR)
118                     Call WriteConsoleMsg(tUser, Mensaje, e_FontTypeNames.FONTTYPE_New_DONADOR)
                    Else
120                     If PersonajeExiste(UserName) Then
122                         Call SetMessageInfoDatabase(UserName, "Mensaje recibido de " & .Name & " [Game Master]: " & vbNewLine & Mensaje & vbNewLine)
                        End If
                    End If

124                 Call WriteConsoleMsg(UserIndex, "Mensaje enviado a " & UserName & ": " & Mensaje, e_FontTypeNames.FONTTYPE_INFO)
126                 Call LogGM(.Name, "Envió mensaje como GM a " & UserName & ": " & Mensaje)

                End If

            End If

        End With
    
        Exit Sub

errHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMensajeUser", Erl)
130

End Sub

Private Sub HandleTraerBoveda(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Ladder
        'Last Modification: 04/jul/2014
        '
        '***************************************************
        On Error GoTo errHandler

100     With UserList(UserIndex)

102         Call UpdateUserHechizos(True, UserIndex, 0)
       
104         Call UpdateUserInv(True, UserIndex, 0)

        End With
    
        Exit Sub

errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTraerBoveda", Erl)
108

End Sub

Private Sub HandleEditChar(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/28/06
        '
        '***************************************************
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
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
        
102         UserName = Replace(Reader.ReadString8(), "+", " ")
        
104         If UCase$(UserName) = "YO" Then
106             tUser = UserIndex
            
            Else
108             tUser = NameIndex(UserName)
            End If
        
110         opcion = Reader.ReadInt8()
112         Arg1 = Reader.ReadString8()
114         Arg2 = Reader.ReadString8()

            ' Si no es GM, no hacemos nada.
116         If Not EsGM(UserIndex) Then Exit Sub
        
            ' Si NO sos Dios o Admin,
118         If (.flags.Privilegios And e_PlayerType.Admin) = 0 Then

                ' Si te editas a vos mismo esta bien ;)
120             If UserIndex <> tUser Then Exit Sub
            
            End If
        
122         Select Case opcion

                Case e_EditOptions.eo_Gold
124
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

126                 If tUser <= 0 Then
128                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
130                     UserList(tUser).Stats.GLD = val(Arg1)
132                     Call WriteUpdateGold(tUser)

                    End If
                
134             Case e_EditOptions.eo_Experience
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

136                 If tUser <= 0 Then
138                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else

140                     If UserList(tUser).Stats.ELV < STAT_MAXELV Then
142                         UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
144                         Call CheckUserLevel(tUser)
146                         Call WriteUpdateExp(tUser)
                            
                        Else
148                         Call WriteConsoleMsg(UserIndex, "El usuario es nivel máximo.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
150             Case e_EditOptions.eo_Body

152                 If tUser <= 0 Then

156                     Call SaveUserBodyDatabase(UserName, val(Arg1))

160                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
162                     Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                    End If
                   
164             Case e_EditOptions.eo_Arma
                
166                 If tUser <= 0 Then
                       
168                     If Database_Enabled Then
                            'Call SaveUserBodyDatabase(UserName, val(Arg1))
                        Else
                            'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                        End If
                    
170                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
172                     Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, val(Arg1), UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    
                    End If
                       
174             Case e_EditOptions.eo_Escudo
                
176                 If tUser <= 0 Then
                       
178                     If Database_Enabled Then
                            'Call SaveUserBodyDatabase(UserName, val(Arg1))
                        Else
                            'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                        End If
                    
180                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
182                     Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, val(Arg1), UserList(tUser).Char.CascoAnim)
                    
                    End If
                       
184             Case e_EditOptions.eo_CASCO
                
186                 If tUser <= 0 Then
                       
188                     If Database_Enabled Then
                            'Call SaveUserBodyDatabase(UserName, val(Arg1))
                        Else
                            'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)
                    
                        End If
                    
190                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
192                     Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, val(Arg1))
                    
                    End If
                       
194             Case e_EditOptions.eo_Particula

196                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                
198                 If Not .flags.Privilegios = Consejero Then
200                     If tUser <= 0 Then

202                         If Database_Enabled Then
                                'Call SaveUserBodyDatabase(UserName, val(Arg1))
                            Else
                                'Call WriteVar(CharPath & UserName & ".chr", "INIT", "Arma", Arg1)

                            End If

204                         Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                        Else
                            'Call ChangeUserChar(tUser, UserList(tUser).Char.Body, UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, val(Arg1))
206                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, val(Arg1), 9999, False))
208                         .Char.ParticulaFx = val(Arg1)
210                         .Char.loops = 9999

                        End If

                    End If
                
212             Case e_EditOptions.eo_Head

214                 If tUser <= 0 Then

218                     Call SaveUserHeadDatabase(UserName, val(Arg1))

222                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
224                     Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                    End If
                
226             Case e_EditOptions.eo_CriminalsKilled
                
228                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                
230                 If tUser <= 0 Then
232                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else

234                     If val(Arg1) > MAXUSERMATADOS Then
236                         UserList(tUser).Faccion.CriminalesMatados = MAXUSERMATADOS
                        Else
238                         UserList(tUser).Faccion.CriminalesMatados = val(Arg1)

                        End If

                    End If
                
240             Case e_EditOptions.eo_CiticensKilled

242                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                
244                 If tUser <= 0 Then
246                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else

248                     If val(Arg1) > MAXUSERMATADOS Then
250                         UserList(tUser).Faccion.ciudadanosMatados = MAXUSERMATADOS
                        Else
252                         UserList(tUser).Faccion.ciudadanosMatados = val(Arg1)

                        End If

                    End If
                
254             Case e_EditOptions.eo_Level
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub

256                 If tUser <= 0 Then
258                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else

260                     If val(Arg1) > STAT_MAXELV Then
262                         Arg1 = CStr(STAT_MAXELV)
264                         Call WriteConsoleMsg(UserIndex, "No podés tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)

                        End If
                        
266                     UserList(tUser).Stats.ELV = val(Arg1)

                    End If
                    
268                 Call WriteUpdateUserStats(UserIndex)
                
270             Case e_EditOptions.eo_Class

272                 If tUser <= 0 Then
274                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else

276                     For LoopC = 1 To NUMCLASES

278                         If Tilde(ListaClases(LoopC)) = Tilde(Arg1) Then Exit For
280                     Next LoopC
                        
282                     If LoopC > NUMCLASES Then
284                         Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", e_FontTypeNames.FONTTYPE_INFO)
                        Else
286                         UserList(tUser).clase = LoopC

                        End If

                    End If
                
288             Case e_EditOptions.eo_Skills
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub

290                 For LoopC = 1 To NUMSKILLS

292                     If Tilde(Replace$(SkillsNames(LoopC), " ", "+")) = Tilde(Arg1) Then Exit For
294                 Next LoopC
                    
296                 If LoopC > NUMSKILLS Then
298                     Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", e_FontTypeNames.FONTTYPE_INFO)
                    Else

300                     If tUser <= 0 Then
                        
304                         Call SaveUserSkillDatabase(UserName, LoopC, val(Arg2))

308                         Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                        Else
310                         UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)

                        End If

                    End If
                
312             Case e_EditOptions.eo_SkillPointsLeft
                
314                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub
                
316                 If tUser <= 0 Then

320                     Call SaveUserSkillsLibres(UserName, val(Arg1))

324                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
326                     UserList(tUser).Stats.SkillPts = val(Arg1)

                    End If
                
328             Case e_EditOptions.eo_Sex

330                 If tUser <= 0 Then
332                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
334                     Arg1 = UCase$(Arg1)

336                     If (Arg1 = "MUJER") Then
338                         UserList(tUser).genero = e_Genero.Mujer
                        
340                     ElseIf (Arg1 = "HOMBRE") Then
342                         UserList(tUser).genero = e_Genero.Hombre

                        End If

                    End If
                
344             Case e_EditOptions.eo_Raza

346                 If tUser <= 0 Then
348                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                        
                    Else
                    
350                     Arg1 = UCase$(Arg1)

352                     If (Arg1 = "HUMANO") Then
354                         UserList(tUser).raza = e_Raza.Humano
                            
356                     ElseIf (Arg1 = "ELFO") Then
358                         UserList(tUser).raza = e_Raza.Elfo
                            
360                     ElseIf (Arg1 = "DROW") Then
362                         UserList(tUser).raza = e_Raza.Drow
                            
364                     ElseIf (Arg1 = "ENANO") Then
366                         UserList(tUser).raza = e_Raza.Enano
                            
368                     ElseIf (Arg1 = "GNOMO") Then
370                         UserList(tUser).raza = e_Raza.Gnomo
                            
372                     ElseIf (Arg1 = "ORCO") Then
374                         UserList(tUser).raza = e_Raza.Orco

                        End If

                    End If
                    
376             Case e_EditOptions.eo_Vida
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

378                 If tUser <= 0 Then
380                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
382                     tmpLong = val(Arg1)
                        
384                     If tmpLong > 0 Then
386                         UserList(tUser).Stats.MaxHp = Min(tmpLong, STAT_MAXHP)
388                         UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MaxHp
                            
390                         Call WriteUpdateUserStats(tUser)

                        End If

                    End If
                    
392             Case e_EditOptions.eo_Mana
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

394                 If tUser <= 0 Then
396                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
398                     tmpLong = val(Arg1)
                        
400                     If tmpLong > 0 Then
402                         UserList(tUser).Stats.MaxMAN = Min(tmpLong, STAT_MAXMP)
404                         UserList(tUser).Stats.MinMAN = UserList(tUser).Stats.MaxMAN
                            
406                         Call WriteUpdateUserStats(tUser)

                        End If

                    End If
                    
408             Case e_EditOptions.eo_Energia

                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

410                 If tUser <= 0 Then
412                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
414                     tmpLong = val(Arg1)
                        
416                     If tmpLong > 0 Then
418                         UserList(tUser).Stats.MaxSta = Min(tmpLong, STAT_MAXSTA)
420                         UserList(tUser).Stats.MinSta = UserList(tUser).Stats.MaxSta
                            
422                         Call WriteUpdateUserStats(tUser)

                        End If

                    End If
                        
424             Case e_EditOptions.eo_MinHP
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

426                 If tUser <= 0 Then
428                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
430                     tmpLong = val(Arg1)
                        
432                     If tmpLong >= 0 Then
434                         UserList(tUser).Stats.MinHp = Min(tmpLong, STAT_MAXHP)
                            
436                         Call WriteUpdateHP(tUser)

                        End If

                    End If
                    
438             Case e_EditOptions.eo_MinMP
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                    
440                 If tUser <= 0 Then
442                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
444                     tmpLong = val(Arg1)
                        
446                     If tmpLong >= 0 Then
448                         UserList(tUser).Stats.MinMAN = Min(tmpLong, STAT_MAXMP)
                            
450                         Call WriteUpdateMana(tUser)

                        End If

                    End If
                    
452             Case e_EditOptions.eo_Hit
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

454                 If tUser <= 0 Then
456                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
458                     tmpLong = val(Arg1)
                        
460                     If tmpLong >= 0 Then
462                         UserList(tUser).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
464                         UserList(tUser).Stats.MinHIT = UserList(tUser).Stats.MaxHit

                        End If

                    End If
                    
466             Case e_EditOptions.eo_MinHit

468                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

                    If tUser <= 0 Then
470                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
472                     tmpLong = val(Arg1)
                        
474                     If tmpLong >= 0 Then
476                         UserList(tUser).Stats.MinHIT = Min(tmpLong, STAT_MAXHIT)

                        End If

                    End If
                    
478             Case e_EditOptions.eo_MaxHit
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

480                 If tUser <= 0 Then
482                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
484                     tmpLong = val(Arg1)
                        
486                     If tmpLong >= 0 Then
488                         UserList(tUser).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)

                        End If

                    End If
                    
490             Case e_EditOptions.eo_Desc

492                 If tUser <= 0 Then
494                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    
496                 ElseIf DescripcionValida(Arg1) Then
498                     UserList(tUser).Desc = Arg1
                        
                    Else
500                     Call WriteConsoleMsg(UserIndex, "Caracteres inválidos en la descripción.", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                    
502             Case e_EditOptions.eo_Intervalo

504                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

506                 If tUser <= 0 Then
508                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
510                     Arg1 = UCase$(Arg1)
                        
512                     tmpLong = val(Arg2)
                        
514                     If tmpLong >= 0 Then
                    
516                         Select Case Arg1

                                Case "USAR"
518                                 UserList(tUser).Intervals.UsarClic = tmpLong
520                                 UserList(tUser).Intervals.UsarU = tmpLong
                                    
522                             Case "USAR_U", "USAR+U", "USAR-U"
524                                 UserList(tUser).Intervals.UsarU = tmpLong
                                    
526                             Case "USAR_CLIC", "USAR+CLIC", "USAR-CLIC", "USAR_CLICK", "USAR+CLICK", "USAR-CLICK"
528                                 UserList(tUser).Intervals.UsarClic = tmpLong
                                    
530                             Case "ARCO", "PROYECTILES"
532                                 UserList(tUser).Intervals.Arco = tmpLong
                                    
534                             Case "GOLPE", "GOLPES", "GOLPEAR"
536                                 UserList(tUser).Intervals.Golpe = tmpLong
                                    
538                             Case "MAGIA", "HECHIZO", "HECHIZOS", "LANZAR"
540                                 UserList(tUser).Intervals.Magia = tmpLong

542                             Case "COMBO"
544                                 UserList(tUser).Intervals.GolpeMagia = tmpLong
546                                 UserList(tUser).Intervals.MagiaGolpe = tmpLong

548                             Case "GOLPE-MAGIA", "GOLPE-HECHIZO"
550                                 UserList(tUser).Intervals.GolpeMagia = tmpLong

552                             Case "MAGIA-GOLPE", "HECHIZO-GOLPE"
554                                 UserList(tUser).Intervals.MagiaGolpe = tmpLong
                                    
556                             Case "GOLPE-USAR"
558                                 UserList(tUser).Intervals.GolpeUsar = tmpLong
                                    
560                             Case "TRABAJAR", "WORK", "TRABAJO"
562                                 UserList(tUser).Intervals.TrabajarConstruir = tmpLong
564                                 UserList(tUser).Intervals.TrabajarExtraer = tmpLong
                                    
566                             Case "TRABAJAR_EXTRAER", "EXTRAER", "TRABAJO_EXTRAER"
568                                 UserList(tUser).Intervals.TrabajarExtraer = tmpLong
                                    
570                             Case "TRABAJAR_CONSTRUIR", "CONSTRUIR", "TRABAJO_CONSTRUIR"
572                                 UserList(tUser).Intervals.TrabajarConstruir = tmpLong
                                    
574                             Case Else
                                    Exit Sub

                            End Select
                            
576                         Call WriteIntervals(tUser)
                            
                        End If

                    End If
                    
578             Case e_EditOptions.eo_Hogar

580                 If tUser <= 0 Then
582                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, e_FontTypeNames.FONTTYPE_INFO)
                    Else
584                     Arg1 = UCase$(Arg1)
                    
586                     Select Case Arg1

                            Case "NIX"
588                             UserList(tUser).Hogar = e_Ciudad.cNix

590                         Case "ULLA", "ULLATHORPE"
592                             UserList(tUser).Hogar = e_Ciudad.cUllathorpe

594                         Case "BANDER", "BANDERBILL"
596                             UserList(tUser).Hogar = e_Ciudad.cBanderbill

598                         Case "LINDOS"
600                             UserList(tUser).Hogar = e_Ciudad.cLindos

602                         Case "ARGHAL"
604                             UserList(tUser).Hogar = e_Ciudad.cArghal

606                         Case "ARKHEIN"
608                             UserList(tUser).Hogar = e_Ciudad.cArkhein

                        End Select

                    End If
                
610             Case Else
                
612                 Call WriteConsoleMsg(UserIndex, "Comando no permitido.", e_FontTypeNames.FONTTYPE_INFO)

            End Select

            'Log it!
614         commandString = "/MOD "
        
616         Select Case opcion

                Case e_EditOptions.eo_Gold
618                 commandString = commandString & "ORO "
            
620             Case e_EditOptions.eo_Experience
622                 commandString = commandString & "EXP "
            
624             Case e_EditOptions.eo_Body
626                 commandString = vbNullString
            
628             Case e_EditOptions.eo_Head
630                 commandString = vbNullString
            
632             Case e_EditOptions.eo_CriminalsKilled
634                 commandString = commandString & "CRI "
            
636             Case e_EditOptions.eo_CiticensKilled
638                 commandString = commandString & "CIU "
            
640             Case e_EditOptions.eo_Level
642                 commandString = commandString & "LEVEL "
            
644             Case e_EditOptions.eo_Class
646                 commandString = commandString & "CLASE "
            
648             Case e_EditOptions.eo_Skills
650                 commandString = commandString & "SKILLS "
            
652             Case e_EditOptions.eo_SkillPointsLeft
654                 commandString = commandString & "SKILLSLIBRES "
                
656             Case e_EditOptions.eo_Sex
658                 commandString = commandString & "SEX "
                
660             Case e_EditOptions.eo_Raza
662                 commandString = commandString & "RAZA "

664             Case e_EditOptions.eo_Vida
666                 commandString = commandString & "VIDA "
                    
668             Case e_EditOptions.eo_Mana
670                 commandString = commandString & "MANA "
                    
672             Case e_EditOptions.eo_Energia
674                 commandString = commandString & "ENERGIA "
                    
676             Case e_EditOptions.eo_MinHP
678                 commandString = commandString & "MINHP "
                    
680             Case e_EditOptions.eo_MinMP
682                 commandString = commandString & "MINMP "
                    
684             Case e_EditOptions.eo_Hit
686                 commandString = commandString & "HIT "
                    
688             Case e_EditOptions.eo_MinHit
690                 commandString = commandString & "MINHIT "
                    
692             Case e_EditOptions.eo_MaxHit
694                 commandString = commandString & "MAXHIT "
                    
696             Case e_EditOptions.eo_Desc
698                 commandString = commandString & "DESC "
                    
700             Case e_EditOptions.eo_Intervalo
702                 commandString = commandString & "INTERVALO "
                    
704             Case e_EditOptions.eo_Hogar
706                 commandString = commandString & "HOGAR "
                
                Case e_EditOptions.eo_CASCO
                    commandString = vbNullString
                   
                Case e_EditOptions.eo_Arma
                    commandString = vbNullString
                    
                Case e_EditOptions.eo_Escudo
                    commandString = vbNullString

708             Case Else
710                 commandString = commandString & "UNKOWN "

            End Select
            
            If commandString <> vbNullString Then
714             Call LogGM(.Name, commandString & Arg1 & " " & Arg2 & " " & UserName)
            End If
            
        End With

        Exit Sub

errHandler:
716     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEditChar", Erl)
718

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
    
            Dim targetName  As String
            Dim TargetIndex As Integer
        
102         targetName = Replace$(Reader.ReadString8(), "+", " ")
104         TargetIndex = NameIndex(targetName)
        
106         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then

                'is the player offline?
108             If TargetIndex > 0 Then

                    'don't allow to retrieve administrator's info
116                 If UserList(TargetIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
118                     Call SendUserStatsTxt(UserIndex, TargetIndex)

                    End If

                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInfo", Erl)
124

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer

102         UserName = Reader.ReadString8()
        
104         If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) <> 0 Then
106             Call LogGM(.Name, "/STAT " & UserName)
            
108             tUser = NameIndex(UserName)
            
110             If tUser > 0 Then

116                 Call SendUserMiniStatsTxt(UserIndex, tUser)

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharStats", Erl)
122

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.Name, "/BAL " & UserName)
            
110             If tUser > 0 Then

116                 Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", e_FontTypeNames.FONTTYPE_TALK)

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharGold", Erl)
122

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.Name, "/INV " & UserName)
            
110             If tUser > 0 Then

116                 Call SendUserInvTxt(UserIndex, tUser)

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInventory", Erl)
122

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.Name, "/BOV " & UserName)
            
110             If tUser <= 0 Then
112                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_TALK)
        
                Else
116                 Call SendUserBovedaTxt(UserIndex, tUser)

                End If
                
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharBank", Erl)
122

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer
            Dim LoopC    As Long
            Dim Message  As String
        
102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.Name, "/STATS " & UserName)
            
110             If tUser <= 0 Then
            
112                 If (InStrB(UserName, "\") <> 0) Then
114                     UserName = Replace(UserName, "\", "")

                    End If

116                 If (InStrB(UserName, "/") <> 0) Then
118                     UserName = Replace(UserName, "/", "")

                    End If
                
120                '  For LoopC = 1 To NUMSKILLS
122                '     Message = Message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
124                 'Next LoopC
                
126                ' Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), e_FontTypeNames.FONTTYPE_INFO)
                
                Else
128                 Call SendUserSkillsTxt(UserIndex, tUser)

                End If
            Else
130             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
132     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharSkills", Erl)
134

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim UserName As String
            Dim tUser    As Integer
            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             If UCase$(UserName) <> "YO" Then
108                 tUser = NameIndex(UserName)
                Else
110                 tUser = UserIndex

                End If
            
112             If tUser <= 0 Then
114                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                Else

116                 With UserList(tUser)

                        'If dead, show him alive (naked).
118                     If .flags.Muerto = 1 Then
120                         .flags.Muerto = 0
                        
                            'Call DarCuerpoDesnudo(tUser)
                        
                            'Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
122                         Call RevivirUsuario(tUser)
                        
124                         Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", e_FontTypeNames.FONTTYPE_INFO)
                        Else
126                         Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", e_FontTypeNames.FONTTYPE_INFO)

                        End If
                    
128                     .Stats.MinHp = .Stats.MaxHp

                    End With
                
                    ' Call WriteHora(tUser)
130                 Call WriteUpdateHP(tUser)
132                 Call ActualizarVelocidadDeUsuario(tUser)
134                 Call LogGM(.Name, "Resucito a " & UserName)

                End If
            Else
136             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
140

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
        Dim priv As e_PlayerType
    
100     With UserList(UserIndex)
         
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        

106         priv = e_PlayerType.Consejero Or e_PlayerType.SemiDios

108         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv Or e_PlayerType.Dios Or e_PlayerType.Admin
        
110         For i = 1 To LastUser

112             If UserList(i).flags.UserLogged Then
114                 If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).Name & ", "

                End If

116         Next i
        
118         If LenB(list) <> 0 Then
120             list = Left$(list, Len(list) - 2)
122             Call WriteConsoleMsg(UserIndex, list & ".", e_FontTypeNames.FONTTYPE_INFO)
            Else
124             Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleOnlineGM_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineGM", Erl)
128
        
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
100     With UserList(UserIndex)
    
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
            Dim LoopC As Long
            Dim list  As String
            Dim priv  As e_PlayerType
        
106         priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios

108         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv + (e_PlayerType.Dios Or e_PlayerType.Admin)
        
110         For LoopC = 1 To LastUser

112             If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = .Pos.Map Then
114                 If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).Name & ", "

                End If

116         Next LoopC
        
118         If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
120         Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleOnlineMap_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineMap", Erl)
124
        
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
100     With UserList(UserIndex)
  
            'Se asegura que el target es un npc
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Validate NPC and make sure player is not dead
106         If (NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        
            Dim priest As t_Npc
108         priest = NpcList(.flags.TargetNPC)

            'Make sure it's close enough
110         If Distancia(.Pos, priest.Pos) > 3 Then
                'Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
112             Call WriteConsoleMsg(UserIndex, "El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If .Faccion.Status = 1 Or .Faccion.ArmadaReal = 1 Then
                'Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
116             Call WriteChatOverHead(UserIndex, "Tu alma ya esta libre de pecados hijo mio.", priest.Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
118         If .Faccion.FuerzasCaos > 0 Then
120             Call WriteChatOverHead(UserIndex, "¡¡Dios no te perdonará mientras seas fiel al Demonio!!", priest.Char.CharIndex, vbWhite)
                Exit Sub

            End If

122         If .GuildIndex <> 0 Then
124             If modGuilds.Alineacion(.GuildIndex) = 1 Then
126                 Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.", priest.Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            End If

128         If .Faccion.ciudadanosMatados > 0 Then
                Dim Donacion As Long
130             Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano

132             Call WriteChatOverHead(UserIndex, "Has matado a ciudadanos inocentes, Dios no puede perdonarte lo que has hecho. " & "Pero si haces una generosa donación de, digamos, " & PonerPuntos(Donacion) & " monedas de oro, tal vez cambie de opinión...", priest.Char.CharIndex, vbWhite)
                Exit Sub

            End If

134         Call WriteChatOverHead(UserIndex, "Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", priest.Char.CharIndex, vbYellow)

136         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, "80", 100, False))
138         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
140         Call VolverCiudadano(UserIndex)

        End With
        
        Exit Sub

HandleForgive_Err:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForgive", Erl)
144
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
            Dim rank     As Integer
        
102         rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
        
104         UserName = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
108             tUser = NameIndex(UserName)
            
110             If tUser <= 0 Then
112                 Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", e_FontTypeNames.FONTTYPE_INFO)
                
                Else

114                 If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
116                     Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
118                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " echo a " & UserName & ".", e_FontTypeNames.FONTTYPE_INFO))
120                     Call CloseSocket(tUser)
122                     Call LogGM(.Name, "Echo a " & UserName)

                    End If

                End If
            Else
124             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKick", Erl)
128

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             tUser = NameIndex(UserName)
            
108             If tUser > 0 Then
 
110                 Call UserDie(tUser)
112                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserList(tUser).Name, e_FontTypeNames.FONTTYPE_EJECUCION))
114                 Call LogGM(.Name, " ejecuto a " & UserName)

                Else
            
116                 Call WriteConsoleMsg(UserIndex, "No está online", e_FontTypeNames.FONTTYPE_INFO)

                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleExecute", Erl)
122

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

        On Error GoTo errHandler

100     With UserList(UserIndex)
  
            Dim UserName As String
            Dim Reason   As String
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call BanPJ(UserIndex, UserName, Reason)
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanChar", Erl)
114

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
102             UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            
106             If Not PersonajeExiste(UserName) Then
108                 Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_INFO)
                Else

110                 If BANCheck(UserName) Then
112                     Call SavePenaDatabase(UserName, .Name & ": UNBAN. " & Date & " " & Time)
114                     Call UnBanDatabase(UserName)

116                     Call LogGM(.Name, "/UNBAN a " & UserName)
118                     Call WriteConsoleMsg(UserIndex, UserName & " desbaneado.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
120                     Call WriteConsoleMsg(UserIndex, UserName & " no esta baneado. Imposible desbanear.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            Else
122             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnbanChar", Erl)
126

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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         If .flags.TargetNPC > 0 Then
        
108             Call DoFollow(.flags.TargetNPC, .Name)
            
110             NpcList(.flags.TargetNPC).flags.Inmovilizado = 0
112             NpcList(.flags.TargetNPC).flags.Paralizado = 0
114             NpcList(.flags.TargetNPC).Contadores.Paralisis = 0

            End If

        End With
        
        Exit Sub

HandleNPCFollow_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNPCFollow", Erl)
118
        
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

    On Error GoTo errHandler

100 With UserList(UserIndex)
    
        Dim UserName As String
        Dim tUser    As Integer
        
102     UserName = Reader.ReadString8()
            
104     If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
106         If LenB(UserName) <> 0 Then
108             tUser = NameIndex(UserName)

110             If tUser <= 0 Then
112                 Call WriteConsoleMsg(UserIndex, "El jugador no está online.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
114         ElseIf .flags.TargetUser > 0 Then
116             tUser = .flags.TargetUser

                ' Mover NPCs
118         ElseIf .flags.TargetNPC > 0 Then

120             If NpcList(.flags.TargetNPC).Pos.Map = .Pos.Map Then
122                 Call WarpNpcChar(.flags.TargetNPC, .Pos.Map, .Pos.x, .Pos.Y + 1, True)
124                 Call WriteConsoleMsg(UserIndex, "Has desplazado a la criatura.", e_FontTypeNames.FONTTYPE_INFO)
                Else
126                 Call WriteConsoleMsg(UserIndex, "Sólo puedes mover NPCs dentro del mismo mapa.", e_FontTypeNames.FONTTYPE_INFO)

                End If

                Exit Sub

            Else
                Exit Sub

            End If

128         If CompararPrivilegiosUser(tUser, UserIndex) > 0 Then
130             Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser).Name & " que quieres traerlo a tu posición.", e_FontTypeNames.FONTTYPE_INFO)
132             Call WriteConsoleMsg(tUser, .Name & " quiere transportarte a su ubicación. Escribe /ira " & .Name & " para ir.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
            Dim NotConsejero As Boolean
134         NotConsejero = (.flags.Privilegios And e_PlayerType.Consejero) = 0
                
            ' Consejeros sólo pueden traer en el mismo mapa
136         If NotConsejero Or .Pos.Map = UserList(tUser).Pos.Map Then
                    
                    ' Si el admin está invisible no mostramos el nombre
138                 If NotConsejero And .flags.AdminInvisible = 1 Then
140                     Call WriteConsoleMsg(tUser, "Te han trasportado.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
142                     Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    'HarThaoS: Si lo sumonean a un mapa interdimensional desde uno no interdimensional me guardo la posición de donde viene.
144                 If EsMapaInterdimensional(.Pos.Map) And Not EsMapaInterdimensional(UserList(tUser).Pos.Map) Then
146                     UserList(tUser).flags.ReturnPos = UserList(tUser).Pos
                    End If
                    
                    

148             Call WarpToLegalPos(tUser, .Pos.Map, .Pos.x, .Pos.Y + 1, True, True)

150             Call WriteConsoleMsg(UserIndex, "Has traído a " & UserList(tUser).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
                    
152             Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.x & " Y:" & .Pos.Y)
                
            End If
        Else
154         Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
        End If

    End With

    Exit Sub
        
errHandler:

156 Call TraceError(Err.Number, Err.Description, "Protocol.HandleSummonChar", Erl)
158

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
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then
                Exit Sub

104         ElseIf .flags.Privilegios And e_PlayerType.Consejero Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
108         ElseIf .flags.Privilegios And e_PlayerType.SemiDios Then
110             Call WriteConsoleMsg(UserIndex, "Servidor » La cantidad de NPCs disponible para tu rango está limitada.", e_FontTypeNames.FONTTYPE_INFO)
            End If

112         Call WriteSpawnList(UserIndex, UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin))
    
        End With
        
        Exit Sub

HandleSpawnListRequest_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnListRequest", Erl)
116
        
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

100     With UserList(UserIndex)

            Dim npc As Integer
102             npc = Reader.ReadInt16()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then
                    If Declaraciones.SpawnList(npc).NpcName <> "Nada" And (Declaraciones.SpawnList(npc).PuedeInvocar Or (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) <> 0) Then
108                     Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
                    End If
                End If
            
110             Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub

HandleSpawnCreature_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnCreature", Erl)
116
        
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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If .flags.TargetNPC = 0 Then Exit Sub
        
108         Call ResetNpcInv(.flags.TargetNPC)
110         Call LogGM(.Name, "/RESETINV " & NpcList(.flags.TargetNPC).Name)

        End With
        
        Exit Sub

HandleResetNPCInventory_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetNPCInventory", Erl)
114
        
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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         Call LimpiezaForzada
            
108         Call WriteConsoleMsg(UserIndex, "Se han limpiado los items del suelo.", e_FontTypeNames.FONTTYPE_INFO)
            
        End With

        Exit Sub

HandleCleanWorld_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanWorld", Erl)
112
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If LenB(Message) <> 0 Then
108                 Call LogGM(.Name, "Mensaje Broadcast:" & Message)
110                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, e_FontTypeNames.FONTTYPE_SERVER))

                End If
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerMessage", Erl)
116

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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
            Dim priv     As e_PlayerType
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             tUser = NameIndex(UserName)
108             Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

110             If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                 priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
                
                Else
114                 priv = e_PlayerType.user

                End If
            
116             If tUser > 0 Then
118                 If UserList(tUser).flags.Privilegios And priv Then
                
120                     Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).IP, e_FontTypeNames.FONTTYPE_INFO)

                        Dim IP    As String
                        Dim lista As String
                        Dim LoopC As Long

122                     IP = UserList(tUser).IP

124                     For LoopC = 1 To LastUser

126                         If UserList(LoopC).IP = IP Then
                        
128                             If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                            
130                                 If UserList(LoopC).flags.Privilegios And priv Then
132                                     lista = lista & UserList(LoopC).Name & ", "
                                    End If

                                End If

                            End If

134                     Next LoopC

136                     If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    
138                     Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, e_FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
            
140                 Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", e_FontTypeNames.FONTTYPE_INFO)

                End If
            Else
142             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNickToIP", Erl)
146

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

100     With UserList(UserIndex)

            Dim IP    As String
            Dim LoopC As Long
            Dim lista As String
            Dim priv  As e_PlayerType
        
102         IP = Reader.ReadInt8() & "."
104         IP = IP & Reader.ReadInt8() & "."
106         IP = IP & Reader.ReadInt8() & "."
108         IP = IP & Reader.ReadInt8()
        
110         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
114         Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & IP)
        
116         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
118             priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
            
            Else
120             priv = e_PlayerType.user

            End If

122         For LoopC = 1 To LastUser

124             If UserList(LoopC).IP = IP Then
            
126                 If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                
128                     If UserList(LoopC).flags.Privilegios And priv Then
130                         lista = lista & UserList(LoopC).Name & ", "
                        End If

                    End If

                End If

132         Next LoopC
        
134         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
136         Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleIPToNick_Err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIPToNick", Erl)
140
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim GuildName As String
            Dim tGuild    As Integer
        
102         GuildName = Reader.ReadString8()
        
104         If (InStrB(GuildName, "+") <> 0) Then
106             GuildName = Replace(GuildName, "+", " ")
            End If
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
110             tGuild = GuildIndex(GuildName)
            
112             If tGuild > 0 Then
114                 Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), e_FontTypeNames.FONTTYPE_GUILDMSG)
                End If
            Else
116             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnlineMembers", Erl)
120

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

100     With UserList(UserIndex)

            Dim Mapa As Integer
            Dim x    As Byte
            Dim Y    As Byte
            Dim Motivo As String
        
102         Mapa = Reader.ReadInt16()
104         x = Reader.ReadInt8()
106         Y = Reader.ReadInt8()
            Motivo = Reader.ReadString8()
        
108         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         Call LogGM(.Name, "/CT " & Mapa & "," & x & "," & Y & "," & Motivo)
        
114         If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, x, Y) Then Exit Sub
        
116         If MapData(.Pos.Map, .Pos.x, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
118         If MapData(.Pos.Map, .Pos.x, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
120         If MapData(Mapa, x, Y).ObjInfo.ObjIndex > 0 Then
122             Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
124         If MapData(Mapa, x, Y).TileExit.Map > 0 Then
126             Call WriteConsoleMsg(UserIndex, "No podés crear un teleport que apunte a la entrada de otro.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Dim Objeto As t_Obj
        
128         Objeto.amount = 1
130         Objeto.ObjIndex = 378
        
132         Call MakeObj(Objeto, .Pos.Map, .Pos.x, .Pos.Y - 1)
        
134         With MapData(.Pos.Map, .Pos.x, .Pos.Y - 1)
136             .TileExit.Map = Mapa
138             .TileExit.x = x
140             .TileExit.Y = Y
            End With

        End With
        
        Exit Sub

HandleTeleportCreate_Err:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportCreate", Erl)
144
        
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
100     With UserList(UserIndex)

            Dim Mapa As Integer
            Dim x    As Byte
            Dim Y    As Byte

            '/dt
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
106         Mapa = .flags.TargetMap
108         x = .flags.TargetX
110         Y = .flags.TargetY
        
112         If Not InMapBounds(Mapa, x, Y) Then Exit Sub
        
114         With MapData(Mapa, x, Y)

                'Si no tengo objeto y no tengo traslado
116             If .ObjInfo.ObjIndex = 0 And .TileExit.Map = 0 Then Exit Sub
                
                'Si no tengo objeto pero tengo traslado
118             If .ObjInfo.ObjIndex = 0 And .TileExit.Map > 0 Then
120                 Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & x & "," & Y)
                
122                 .TileExit.Map = 0
124                 .TileExit.x = 0
126                 .TileExit.Y = 0
                
                    'si tengo objeto y traslado
128             ElseIf .ObjInfo.ObjIndex > 0 And ObjData(.ObjInfo.ObjIndex).OBJType = e_OBJType.otTeleport Then
130                 Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & x & "," & Y)
                
132                 Call EraseObj(.ObjInfo.amount, Mapa, x, Y)
                
134                 If MapData(.TileExit.Map, .TileExit.x, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
136                     Call EraseObj(1, .TileExit.Map, .TileExit.x, .TileExit.Y)

                    End If
                
138                 .TileExit.Map = 0
140                 .TileExit.x = 0
142                 .TileExit.Y = 0

                End If

            End With

        End With
        
        Exit Sub

HandleTeleportDestroy_Err:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportDestroy", Erl)
146
        
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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         Call LogGM(.Name, "/LLUVIA")
        
108         Lloviendo = Not Lloviendo
110         Nebando = Not Nebando
        
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
114         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

116         If Lloviendo Then
118             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
120             Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HF5D3F3, 250)) 'Rayo
122             Call ApagarFogatas

            End If

        End With
        
        Exit Sub

HandleRainToggle_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
126
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim tUser As Integer
            Dim Desc  As String
        
102         Desc = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
        
106             tUser = .flags.TargetUser

108             If tUser > 0 Then
110                 UserList(tUser).DescRM = Desc
                
                Else
112                 Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetCharDescription", Erl)
116

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

100     With UserList(UserIndex)
    
            Dim midiID As Byte
            Dim Mapa   As Integer
        
102         midiID = Reader.ReadInt8
104         Mapa = Reader.ReadInt16
        
            'Solo dioses, admins y RMS
106         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then

                'Si el mapa no fue enviado tomo el actual
108             If Not InMapBounds(Mapa, 50, 50) Then
110                 Mapa = .Pos.Map
                End If
        
112             If midiID = 0 Then
                    'Ponemos el default del mapa
114                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).music_numberLow))
                
                Else
                    'Ponemos el pedido por el GM
116                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))

                End If

            End If

        End With
        
        Exit Sub

HanldeForceMIDIToMap_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HanldeForceMIDIToMap", Erl)
120
        
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
    
100     With UserList(UserIndex)

            Dim waveID As Byte
            Dim Mapa   As Integer
            Dim x      As Byte
            Dim Y      As Byte
        
102         waveID = Reader.ReadInt8()
104         Mapa = Reader.ReadInt16()
106         x = Reader.ReadInt8()
108         Y = Reader.ReadInt8()
        
            'Solo dioses, admins y RMS
110         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then

                'Si el mapa no fue enviado tomo el actual
112             If Not InMapBounds(Mapa, x, Y) Then
            
114                 Mapa = .Pos.Map
116                 x = .Pos.x
118                 Y = .Pos.Y

                End If
            
                'Ponemos el pedido por el GM
120             Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, x, Y))

            End If

        End With
        
        Exit Sub

HandleForceWAVEToMap_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEToMap", Erl)
124
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
106             Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & Message, e_FontTypeNames.FONTTYPE_TALK))
            End If

        End With

        Exit Sub

errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyMessage", Erl)
110

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
106             Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, e_FontTypeNames.FONTTYPE_TALK))
            End If

        End With

        Exit Sub

errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionMessage", Erl)
110

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
106             Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, e_FontTypeNames.FONTTYPE_TALK))
            End If

        End With

        Exit Sub

errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCitizenMessage", Erl)
110

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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
106             Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, e_FontTypeNames.FONTTYPE_TALK))
            End If

        End With

        Exit Sub

errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCriminalMessage", Erl)
110

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
        
            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then

                'Asegurarse haya un NPC seleccionado
106             If .flags.TargetNPC > 0 Then
108                 Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite))
                
                Else
110                 Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With

        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalkAsNPC", Erl)
114

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
100     With UserList(UserIndex)
  
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim x As Long
            Dim Y As Long
        
106         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
108             For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1

110                 If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                
112                     If MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex > 0 Then
                    
114                         If ItemNoEsDeMapa(MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex) Then
116                             Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, x, Y)
                            End If

                        End If

                    End If

118             Next x
120         Next Y
        
122         Call LogGM(UserList(UserIndex).Name, "/MASSDEST")

        End With
        
        Exit Sub

HandleDestroyAllItemsInArea_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyAllItemsInArea", Erl)
126
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)
    
            Dim UserName As String
            Dim tUser    As Integer
            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
            
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", e_FontTypeNames.FONTTYPE_CONSEJO))

114                 With UserList(tUser)

116                     If .flags.Privilegios And e_PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - e_PlayerType.ChaosCouncil
118                     If Not .flags.Privilegios And e_PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + e_PlayerType.RoyalCouncil
                    
120                     Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y, False)

                    End With

                End If

            End If

        End With

        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptRoyalCouncilMember", Erl)
124

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
            
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legión Oscura.", e_FontTypeNames.FONTTYPE_CONSEJO))
                
114                 With UserList(tUser)

116                     If .flags.Privilegios And e_PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - e_PlayerType.RoyalCouncil
118                     If Not .flags.Privilegios And e_PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + e_PlayerType.ChaosCouncil

120                     Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y, False)

                    End With

                End If

            End If

        End With

        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptChaosCouncilMember", Erl)
124

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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim tObj  As Integer
            Dim lista As String
            Dim x     As Long
            Dim Y     As Long
        
106         For x = 5 To 95
108             For Y = 5 To 95
110                 tObj = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex

112                 If tObj > 0 Then
                
114                     If ObjData(tObj).OBJType <> e_OBJType.otArboles Then
116                         Call WriteConsoleMsg(UserIndex, "(" & x & "," & Y & ") " & ObjData(tObj).Name, e_FontTypeNames.FONTTYPE_INFO)
                        End If

                    End If

118             Next Y
120         Next x

        End With
        
        Exit Sub

HandleItemsInTheFloor_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleItemsInTheFloor", Erl)
124
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)
                
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(UserName)

                'para deteccion de aoice
108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Offline", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
112                 Call WriteDumb(tUser)

                End If

            End If

        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumb", Erl)
116

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
        On Error GoTo errHandler

100     With UserList(UserIndex)
    
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(UserName)

                'para deteccion de aoice
108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Offline", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
112                 Call WriteDumbNoMore(tUser)

                End If

            End If

        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumbNoMore", Erl)
116

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
  
        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 If PersonajeExiste(UserName) Then
112                     Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos", e_FontTypeNames.FONTTYPE_INFO)

116                     Call EcharConsejoDatabase(UserName)

                    Else
122                     Call WriteConsoleMsg(UserIndex, "No existe el personaje.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

124                 With UserList(tUser)

126                     If .flags.Privilegios And e_PlayerType.RoyalCouncil Then
128                         Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", e_FontTypeNames.FONTTYPE_TALK)
130                         .flags.Privilegios = .flags.Privilegios - e_PlayerType.RoyalCouncil
                        
132                         Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y)
134                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", e_FontTypeNames.FONTTYPE_CONSEJO))

                        End If
                    
136                     If .flags.Privilegios And e_PlayerType.ChaosCouncil Then
138                         Call WriteConsoleMsg(tUser, "Has sido echado del consejo de la Legión Oscura", e_FontTypeNames.FONTTYPE_TALK)
140                         .flags.Privilegios = .flags.Privilegios - e_PlayerType.ChaosCouncil
                        
142                         Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y)
144                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legión Oscura", e_FontTypeNames.FONTTYPE_CONSEJO))

                        End If

                    End With

                End If

            End If

        End With

        Exit Sub

errHandler:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilKick", Erl)
148

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

100     With UserList(UserIndex)
 
            Dim tTrigger As Byte
            Dim tLog     As String
        
102         tTrigger = Reader.ReadInt8()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
106         If tTrigger >= 0 Then
        
108             MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = tTrigger
            
110             tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.x & "," & .Pos.Y
            
112             Call LogGM(.Name, tLog)
            
114             Call WriteConsoleMsg(UserIndex, tLog, e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleSetTrigger_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetTrigger", Erl)
118
        
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
    
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         tTrigger = MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger
        
106         Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.x & "," & .Pos.Y & ". Era " & tTrigger)
        
108         Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.x & ", " & .Pos.Y, e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleAskTrigger_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAskTrigger", Erl)
112
        
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBannedIPList_Err

100 With UserList(UserIndex)
102     If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub

        Dim lista As String
        Dim LoopC As Long

104     Call LogGM(.Name, "/BANIPLIST")
        
106         For LoopC = 1 To IP_Blacklist.Count
108             lista = lista & IP_Blacklist.Item(LoopC) & ", "
110         Next LoopC

        
112     If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
114     Call WriteConsoleMsg(UserIndex, lista, e_FontTypeNames.FONTTYPE_INFO)

    End With
        
    Exit Sub

HandleBannedIPList_Err:
116 Call TraceError(Err.Number, Err.Description, "Protocol.HandleBannedIPList", Erl)
118
        
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
        
    On Error GoTo HandleBannedIPReload_Err

100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub

104         Call CargarListaNegraUsuarios
            
106         Call WriteConsoleMsg(UserIndex, "Lista de IPs recargada.", e_FontTypeNames.FONTTYPE_INFO)
            
    End With
        
    Exit Sub

HandleBannedIPReload_Err:
108 Call TraceError(Err.Number, Err.Description, "Protocol.HandleBannedIPReload", Erl)
110
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim GuildName   As String
            Dim cantMembers As Integer
            Dim LoopC       As Long
            Dim member      As String
            Dim Count       As Byte
            Dim tIndex      As Integer
            Dim tFile       As String
        
102         GuildName = Reader.ReadString8()
        
104         If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
106             tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
108             If Not FileExist(tFile) Then
110                 Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, e_FontTypeNames.FONTTYPE_INFO)
                
                Else
            
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " banned al clan " & UCase$(GuildName), e_FontTypeNames.FONTTYPE_FIGHT))
                
                    'baneamos a los miembros
114                 Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
116                 cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
118                 For LoopC = 1 To cantMembers
                        'member es la victima
120                     member = GetVar(tFile, "Members", "Member" & LoopC)
                    
122                     Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
124                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", e_FontTypeNames.FONTTYPE_FIGHT))
                    
126                     tIndex = NameIndex(member)

128                     If tIndex > 0 Then
                            'esta online
130                         UserList(tIndex).flags.Ban = 1
132                         Call CloseSocket(tIndex)

                        End If
                    
136                     Call SaveBanDatabase(member, .Name & " - BAN AL CLAN: " & GuildName & ". " & Date & " " & Time, .Name)


150                 Next LoopC

                End If

            End If

        End With

        Exit Sub

errHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildBan", Erl)
154

End Sub

''
' Handles the "BanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
        Dim tUser As Integer
        Dim bannedIP As String
        
100     With UserList(UserIndex)
        
102         Dim NickOrIP As String: NickOrIP = Reader.ReadString8()
104         Dim Reason As String: Reason = Reader.ReadString8()
        
            ' Si el 4to caracter es un ".", de "XXX.XXX.XXX.XXX", entonces es IP.
106         If mid$(NickOrIP, 4, 1) = "." Then
            
                ' Me fijo que tenga formato valido
108             If IsValidIPAddress(NickOrIP) Then
110                 bannedIP = NickOrIP
                Else
112                 Call WriteConsoleMsg(UserIndex, "La IP " & NickOrIP & " no tiene un formato válido.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
            Else ' Es un Nick
        
114             tUser = NameIndex(NickOrIP)
                
116             If tUser <= 0 Then
118                 Call WriteConsoleMsg(UserIndex, "El personaje no está online.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
120                 bannedIP = UserList(tUser).IP
                End If
            
            End If
         
122         If LenB(bannedIP) = 0 Then Exit Sub
        
124         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
126             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
      
128         If IP_Blacklist.Exists(bannedIP) Then
130             Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista negra de IPs.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
132         Call BanearIP(UserIndex, NickOrIP, bannedIP)
        
134         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, e_FontTypeNames.FONTTYPE_FIGHT))
        
        
            'Find every player with that ip and ban him!
            Dim i As Long
136         For i = 1 To LastUser

138             If UserList(i).ConnIDValida Then
            
140                 If UserList(i).IP = bannedIP Then
                
142                     Call WriteCerrarleCliente(i)
144                     Call CloseSocket(i)
                    
                    End If
                
                End If

146         Next i

        End With

        Exit Sub

errHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanIP", Erl)
150

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUnbanIP_Err

100     With UserList(UserIndex)
        
            Dim bannedIP As String
        
102         bannedIP = Reader.ReadInt8() & "."
104         bannedIP = bannedIP & Reader.ReadInt8() & "."
106         bannedIP = bannedIP & Reader.ReadInt8() & "."
108         bannedIP = bannedIP & Reader.ReadInt8()
        
110         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub
        
112         If IP_Blacklist.Exists(bannedIP) Then
114             Call DesbanearIP(bannedIP, UserIndex)
116             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", e_FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleUnbanIP_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnbanIP", Erl)
122
        
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
100     With UserList(UserIndex)

            Dim tObj    As Integer
            Dim Cuantos As Integer
        
102         tObj = Reader.ReadInt16()
104         Cuantos = Reader.ReadInt16()
    
            ' Si es usuario, lo sacamos cagando.
106         If Not EsGM(UserIndex) Or (.flags.Privilegios And e_PlayerType.Consejero) Then Exit Sub
        
            ' Si es Semi-Dios, dejamos crear un item siempre y cuando pueda estar en el inventario.
108         If (.flags.Privilegios And e_PlayerType.SemiDios) <> 0 And ObjData(tObj).Agarrable = 1 Then Exit Sub

            ' Si hace mas de 10000, lo sacamos cagando.
110         If Cuantos > MAX_INVENTORY_OBJS Then
112             Call WriteConsoleMsg(UserIndex, "Solo podés crear hasta " & CStr(MAX_INVENTORY_OBJS) & " unidades", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If
        
            ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
114         If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
            ' El nombre del objeto es nulo?
116         If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
            Dim Objeto As t_Obj
118         Objeto.amount = Cuantos
120         Objeto.ObjIndex = tObj

            ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demAs t_Objs. que no deberian estar en el inventario)
            '   0 = SI
            '   1 = NO
122         If ObjData(tObj).Agarrable = 0 Then
            
                ' Trato de meterlo en el inventario.
124             If MeterItemEnInventario(UserIndex, Objeto) Then
126                 Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.amount & " unidades de " & ObjData(tObj).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
            
                Else

128                 Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", e_FontTypeNames.FONTTYPE_INFO)
                
                    ' Si no hay espacio y es Dios o Admin, lo tiro al piso.
130                 If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
132                     Call TirarItemAlPiso(.Pos, Objeto)
134                     Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", e_FontTypeNames.FONTTYPE_GUILD)

                    End If
                
                End If
        
            Else
        
                ' Crear el item NO AGARRARBLE y tirarlo al piso.
                ' Si no hay espacio y es Dios o Admin, lo tiro al piso.
136             If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
138                 Call TirarItemAlPiso(.Pos, Objeto)
140                 Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", e_FontTypeNames.FONTTYPE_GUILD)

                End If

            End If
        
            ' Lo registro en los logs.
142         Call LogGM(.Name, "/CI: " & tObj & " Cantidad : " & Cuantos)

        End With
        
        Exit Sub

HandleCreateItem_Err:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateItem", Erl)
146
        
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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         If MapData(.Pos.Map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
108         Call LogGM(.Name, "/DEST")

110         Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, .Pos.x, .Pos.Y)

        End With
        
        Exit Sub

HandleDestroyItems_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyItems", Erl)
114
        
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
        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")

                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")

                End If

114             tUser = NameIndex(UserName)
            
116             Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
118             If tUser > 0 Then
120                 UserList(tUser).Faccion.FuerzasCaos = 0
122                 UserList(tUser).Faccion.Reenlistadas = 200
124                 Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
126                 Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", e_FontTypeNames.FONTTYPE_FIGHT)
                
                Else

128                 If PersonajeExiste(UserName) Then
132                     Call EcharLegionDatabase(UserName)
 
140                     Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
                    Else
142                     Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End With

        Exit Sub

errHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionKick", Erl)
146

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
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
        
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")
                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")
                End If

114             tUser = NameIndex(UserName)
            
116             Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            
118             If tUser > 0 Then
            
120                 UserList(tUser).Faccion.ArmadaReal = 0
122                 UserList(tUser).Faccion.Reenlistadas = 200
                
124                 Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
126                 Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", e_FontTypeNames.FONTTYPE_FIGHT)
                
                Else

128                 If PersonajeExiste(UserName) Then

132                     Call EcharArmadaDatabase(UserName)

140                     Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
142                     Call WriteConsoleMsg(UserIndex, UserName & " inexistente.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End With

        Exit Sub

errHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyKick", Erl)
146

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

100     With UserList(UserIndex)

            Dim midiID As Byte
102             midiID = Reader.ReadInt8()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica: " & midiID, e_FontTypeNames.FONTTYPE_SERVER))
110         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))

        End With
        
        Exit Sub

HandleForceMIDIAll_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceMIDIAll", Erl)
114
        
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
    
100     With UserList(UserIndex)

            Dim waveID As Byte
102             waveID = Reader.ReadInt8()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))

        End With
        
        Exit Sub

HandleForceWAVEAll_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEAll", Erl)
112
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName   As String
            Dim punishment As Byte
            Dim NewText    As String
        
102         UserName = Reader.ReadString8()
104         punishment = Reader.ReadInt8
106         NewText = Reader.ReadString8()
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
        
110             If LenB(UserName) = 0 Then
112                 Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", e_FontTypeNames.FONTTYPE_INFO)
                
                Else

114                 If (InStrB(UserName, "\") <> 0) Then
116                     UserName = Replace(UserName, "\", "")

                    End If

118                 If (InStrB(UserName, "/") <> 0) Then
120                     UserName = Replace(UserName, "/", "")

                    End If
                
122                 If PersonajeExiste(UserName) Then
124                     Call LogGM(.Name, "Borro la pena " & punishment & " de " & UserName & " y la cambió por: " & NewText)

128                     Call CambiarPenaDatabase(UserName, punishment, .Name & ": <" & NewText & "> " & Date & " " & Time)

132                     Call WriteConsoleMsg(UserIndex, "Pena Modificada.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End With

        Exit Sub

errHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemovePunishment", Erl)
136

End Sub

''
' Handles the "Tile_BlockedToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTile_BlockedToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleTile_BlockedToggle_Err

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        

106         Call LogGM(.Name, "/BLOQ")
        
108         If MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = 0 Then
110             MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
            
            Else
112             MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = 0

            End If
        
114         Call Bloquear(True, .Pos.Map, .Pos.x, .Pos.Y, IIf(MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked > 0, e_Block.ALL_SIDES, 0))

        End With
        
        Exit Sub

HandleTile_BlockedToggle_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTile_BlockedToggle", Erl)
118
        
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
100     With UserList(UserIndex)

            If Not EsGM(UserIndex) Then Exit Sub

102         If .flags.Privilegios And e_PlayerType.Consejero Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         If .flags.TargetNPC = 0 Then Exit Sub
        
108         Call QuitarNPC(.flags.TargetNPC)
        
110         Call LogGM(.Name, "/MATA " & NpcList(.flags.TargetNPC).Name)

        End With
        
        Exit Sub

HandleKillNPCNoRespawn_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPCNoRespawn", Erl)
114
        
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
        
        On Error GoTo HandleKillAllNearbyNPCs_Err

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/07/2021
        'ReyarB
        '***************************************************
100     With UserList(UserIndex)

            If Not EsGM(UserIndex) Then Exit Sub
        
102         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            
            'Si está en el mapa pretoriano, me aseguro de que los saque correctamente antes que nada.
106         If .Pos.Map = MAPA_PRETORIANO Then Call EliminarPretorianos(MAPA_PRETORIANO)

            Dim x As Long
            Dim Y As Long
        
108         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
110             For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1

112                 If x > 0 And Y > 0 And x < 101 And Y < 101 Then

114                     If MapData(.Pos.Map, x, Y).NpcIndex > 0 Then
                    
116                         Call QuitarNPC(MapData(.Pos.Map, x, Y).NpcIndex)

                        End If

                    End If

118             Next x
120         Next Y

122         Call LogGM(.Name, "/MASSKILL")

        End With
        
        Exit Sub

HandleKillAllNearbyNPCs_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillAllNearbyNPCs", Erl)
126
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName   As String
            Dim lista      As String
            Dim LoopC      As Byte
            Dim priv       As Integer
            Dim validCheck As Boolean
        
102         priv = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
104         UserName = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then

                'Handle special chars
108             If (InStrB(UserName, "\") <> 0) Then
110                 UserName = Replace(UserName, "\", "")
                End If

112             If (InStrB(UserName, "\") <> 0) Then
114                 UserName = Replace(UserName, "/", "")
                End If

116             If (InStrB(UserName, "+") <> 0) Then
118                 UserName = Replace(UserName, "+", " ")
                End If
            
                'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
120             If NameIndex(UserName) > 0 Then
122                 validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0
                
                Else
124                 validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0

                End If
            
126             If validCheck Then
128                 Call LogGM(.Name, "/LASTIP " & UserName)
                
130                 If FileExist(CharPath & UserName & ".chr", vbNormal) Then
132                     'lista = "Las ultimas IPs con las que " & UserName & " se conectí son:"

134                     'For LoopC = 1 To 5
136                      '   lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
138                     'Next LoopC

140                     'Call WriteConsoleMsg(UserIndex, lista, e_FontTypeNames.FONTTYPE_INFO)
                    
                    Else
142                     Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
144                 Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", e_FontTypeNames.FONTTYPE_INFO)

                End If
            Else
146             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLastIP", Erl)
150

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

100     With UserList(UserIndex)

            Dim Color As Long
102             Color = RGB(Reader.ReadInt8(), Reader.ReadInt8(), Reader.ReadInt8())
        
104         If EsGM(UserIndex) Then
106             .flags.ChatColor = Color
            End If

        End With
        
        Exit Sub

HandleChatColor_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChatColor", Erl)
110
        
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
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
104             .flags.AdminPerseguible = Not .flags.AdminPerseguible
            End If

        End With
        
        Exit Sub

HandleIgnored_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIgnored", Erl)
108
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            'Reads the UserName and Slot Packets
            Dim UserName As String
            Dim Slot     As Byte
            Dim tIndex   As Integer
        
102         UserName = Reader.ReadString8() 'Que UserName?
104         Slot = Reader.ReadInt8() 'Que Slot?
106         tIndex = NameIndex(UserName)  'Que user index?

108         If Not EsGM(UserIndex) Then Exit Sub
        
110         Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & UserName)
           
112         If tIndex > 0 Then
114             If Slot > 0 And Slot <= UserList(UserIndex).CurrentInventorySlots Then
            
116                 If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
118                     Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).amount, e_FontTypeNames.FONTTYPE_INFO)
                    Else
120                     Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
122                 Call WriteConsoleMsg(UserIndex, "Slot Invílido.", e_FontTypeNames.FONTTYPE_TALK)

                End If

            Else
124             Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_TALK)

            End If

        End With
    
        Exit Sub

errHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCheckSlot", Erl)
128

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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'time and Time BUG!
106         Call LogGM(.Name, .Name & " reinicio el mundo")
        
108         Call ReiniciarServidor(True)

        End With
        
        Exit Sub

HandleRestart_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRestart", Erl)
112
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha recargado a los objetos.")
        
106         Call LoadOBJData
108         Call LoadPesca
110         Call LoadRecursosEspeciales
112         Call WriteConsoleMsg(UserIndex, "Obj.dat recargado exitosamente.", e_FontTypeNames.FONTTYPE_SERVER)

        End With
        
        Exit Sub

HandleReloadObjects_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadObjects", Erl)
116
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
106         Call CargarHechizos

        End With
        
        Exit Sub

HandleReloadSpells_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadSpells", Erl)
110
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
106         Call LoadSini

        End With
        
        Exit Sub

HandleReloadServerIni_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadServerIni", Erl)
110
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
         
104         Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
106         Call CargaNpcsDat
    
108         Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado exitosamente.", e_FontTypeNames.FONTTYPE_SERVER)

        End With
        
        Exit Sub

HandleReloadNPCs_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadNPCs", Erl)
112
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
106         Call EcharPjsNoPrivilegiados

        End With
        
        Exit Sub

HandleKickAllChars_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKickAllChars", Erl)
110
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

106         HoraMundo = GetTickCount()

108         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With
        
        Exit Sub

HandleNight_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNight", Erl)
112
        
End Sub

''
' Handle the "Day" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDay(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDay_Err

100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

106         HoraMundo = GetTickCount() - DuracionDia \ 2

108         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With
        
        Exit Sub

HandleDay_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDay", Erl)
112
        
End Sub

''
' Handle the "SetTime" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetTime(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSetTime_Err

100     With UserList(UserIndex)
        
        

            Dim HoraDia As Long
102         HoraDia = Reader.ReadInt32
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

108         HoraMundo = GetTickCount() - HoraDia
            
110         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With
        
        Exit Sub

HandleSetTime_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetTime", Erl)
114
        
End Sub

Public Sub HandleDonateGold(ByVal UserIndex As Integer)
        
        On Error GoTo handle

100     With UserList(UserIndex)
        
        

            Dim Oro As Long
102         Oro = Reader.ReadInt32

104         If Oro <= 0 Then Exit Sub

            'Se asegura que el target es un npc
106         If .flags.TargetNPC = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Dim priest As t_Npc
110         priest = NpcList(.flags.TargetNPC)

            'Validate NPC is an actual priest and the player is not dead
112         If (priest.NPCtype <> e_NPCType.Revividor And (priest.NPCtype <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub

            'Make sure it's close enough
114         If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 3 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

118         If .Faccion.Status = 1 Or .Faccion.ArmadaReal = 1 Or .Faccion.FuerzasCaos > 0 Or .Faccion.ciudadanosMatados = 0 Then
120             Call WriteChatOverHead(UserIndex, "No puedo aceptar tu donación en este momento...", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

122         If .GuildIndex <> 0 Then
124             If modGuilds.Alineacion(.GuildIndex) = 1 Then
126                 Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... no puedo aceptar tu donación.", priest.Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            End If

128         If .Stats.GLD < Oro Then
130             Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Dim Donacion As Long
132         Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano

134         If Oro < Donacion Then
136             Call WriteChatOverHead(UserIndex, "Dios no puede perdonarte si eres una persona avara.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

138         .Stats.GLD = .Stats.GLD - Oro

140         Call WriteUpdateGold(UserIndex)

142         Call WriteConsoleMsg(UserIndex, "Has donado " & PonerPuntos(Oro) & " monedas de oro.", e_FontTypeNames.FONTTYPE_INFO)

144         Call WriteChatOverHead(UserIndex, "¡Gracias por tu generosa donación! Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbYellow)

146         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, "80", 100, False))
148         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
150         Call VolverCiudadano(UserIndex)
    
        End With
        
        Exit Sub

handle:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDonateGold", Erl)
154
        
End Sub

Public Sub HandlePromedio(ByVal UserIndex As Integer)
        
        On Error GoTo handle

100     With UserList(UserIndex)

102         Call WriteConsoleMsg(UserIndex, ListaClases(.clase) & " " & ListaRazas(.raza) & " nivel " & .Stats.ELV & ".", FONTTYPE_INFOBOLD)
            
            Dim Promedio As Double, Vida As Long
        
104         Promedio = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
106         Vida = 18 + ModRaza(.raza).Constitucion + Promedio * (.Stats.ELV - 1)

108         Call WriteConsoleMsg(UserIndex, "Vida esperada: " & Vida & ". Promedio: " & Promedio, FONTTYPE_INFOBOLD)

110         Promedio = CalcularPromedioVida(UserIndex)

            Dim Diff As Long, Color As e_FontTypeNames, Signo As String
            
112         Diff = .Stats.MaxHp - Vida
            
114         If Diff < 0 Then
116             Color = FONTTYPE_PROMEDIO_MENOR
118             Signo = "-"

120         ElseIf Diff > 0 Then
122             Color = FONTTYPE_PROMEDIO_MAYOR
124             Signo = "+"

            Else
126             Color = FONTTYPE_PROMEDIO_IGUAL
128             Signo = "+"
                
            End If

130         Call WriteConsoleMsg(UserIndex, "Vida actual: " & .Stats.MaxHp & " (" & Signo & Abs(Diff) & "). Promedio: " & Round(Promedio, 2), Color)

        End With
        
        Exit Sub

handle:
132     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePromedio", Erl)
134
        
End Sub

Public Sub HandleGiveItem(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim ObjIndex As Integer
            Dim Cantidad As Integer
            Dim Motivo   As String
            Dim tIndex   As Integer
        
102         UserName = Reader.ReadString8()
104         ObjIndex = Reader.ReadInt16()
106         Cantidad = Reader.ReadInt16()
108         Motivo = Reader.ReadString8()
        
110         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then

112             If ObjData(ObjIndex).Agarrable = 1 Then Exit Sub

114             If Cantidad > MAX_INVENTORY_OBJS Then Cantidad = MAX_INVENTORY_OBJS

                ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
116             If ObjIndex < 1 Or ObjIndex > NumObjDatas Then Exit Sub
            
                ' El nombre del objeto es nulo?
118             If LenB(ObjData(ObjIndex).Name) = 0 Then Exit Sub

                ' Está online?
120             tIndex = NameIndex(UserName)

122             If tIndex = 0 Then
124                 Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no está conectado.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                Dim Objeto As t_Obj
126             Objeto.amount = Cantidad
128             Objeto.ObjIndex = ObjIndex

                ' Trato de meterlo en el inventario.
130             If MeterItemEnInventario(tIndex, Objeto) Then
132                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha otorgado a " & UserList(tIndex).Name & " " & Cantidad & " " & ObjData(ObjIndex).Name & ": " & Motivo, e_FontTypeNames.FONTTYPE_ROSA))
                Else
134                 Call WriteConsoleMsg(UserIndex, "El usuario no tiene espacio en el inventario.", e_FontTypeNames.FONTTYPE_INFO)

                End If

                ' Lo registro en los logs.
136             Call LogGM(.Name, "/DAR " & UserName & " - Item: " & ObjData(ObjIndex).Name & "(" & ObjIndex & ") Cantidad : " & Cantidad)
138             Call LogPremios(.Name, UserName, ObjIndex, Cantidad, Motivo)
            Else
140             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo, debes pedir a un Dios que lo de.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGiveItem", Erl)
144
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
106         Call frmMain.mnuMostrar_Click

        End With
        
        Exit Sub

HandleShowServerForm_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowServerForm", Erl)
110
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha borrado los SOS")
        
106         Call Ayuda.Reset

        End With
        
        Exit Sub

HandleCleanSOS_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanSOS", Erl)
110
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         Call LogGM(.Name, .Name & " ha guardado todos los chars")
        
108         Call GuardarUsuarios

        End With
        
        Exit Sub

HandleSaveChars_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSaveChars", Erl)
112
        
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

100     With UserList(UserIndex)

            Dim doTheBackUp As Boolean
        
102         doTheBackUp = Reader.ReadBool()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
106         Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp")
        
            'Change the boolean to byte in a fast way
108         If doTheBackUp Then
110             MapInfo(.Pos.Map).backup_mode = 1
            
            Else
112             MapInfo(.Pos.Map).backup_mode = 0

            End If
        
            'Change the boolean to string in a fast way
114         Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).backup_mode)
        
116         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).backup_mode, e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleChangeMapInfoBackup_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoBackup", Erl)
120
        
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
100     With UserList(UserIndex)

            Dim isMapPk As Boolean
        
102         isMapPk = Reader.ReadBool()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es seguro el mapa.")
        
110         MapInfo(.Pos.Map).Seguro = IIf(isMapPk, 1, 0)

112         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Seguro: " & MapInfo(.Pos.Map).Seguro, e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleChangeMapInfoPK_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoPK", Erl)
116
        
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
        On Error GoTo errHandler

        Dim tStr As String
    
100     With UserList(UserIndex)

102         tStr = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) <> 0 Then

106             Select Case UCase$(tStr)
                
                    Case "NEWBIE"
108                     MapInfo(.Pos.Map).Newbie = Not MapInfo(.Pos.Map).Newbie
110                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": Newbie = " & MapInfo(.Pos.Map).Newbie, e_FontTypeNames.FONTTYPE_INFO)
112                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": Newbie = " & MapInfo(.Pos.Map).Newbie)
                        
114                 Case "SINMAGIA"
116                     MapInfo(.Pos.Map).SinMagia = Not MapInfo(.Pos.Map).SinMagia
118                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": SinMagia = " & MapInfo(.Pos.Map).SinMagia, e_FontTypeNames.FONTTYPE_INFO)
120                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinMagia = " & MapInfo(.Pos.Map).SinMagia)
                        
122                 Case "NOPKS"
124                     MapInfo(.Pos.Map).NoPKs = Not MapInfo(.Pos.Map).NoPKs
126                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": NoPKs = " & MapInfo(.Pos.Map).NoPKs, e_FontTypeNames.FONTTYPE_INFO)
128                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoPKs = " & MapInfo(.Pos.Map).NoPKs)
                        
130                 Case "NOCIUD"
132                     MapInfo(.Pos.Map).NoCiudadanos = Not MapInfo(.Pos.Map).NoCiudadanos
134                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos, e_FontTypeNames.FONTTYPE_INFO)
136                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos)
                        
138                 Case "SININVI"
140                     MapInfo(.Pos.Map).SinInviOcul = Not MapInfo(.Pos.Map).SinInviOcul
142                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & ": SinInvi = " & MapInfo(.Pos.Map).SinInviOcul, e_FontTypeNames.FONTTYPE_INFO)
144                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinInvi = " & MapInfo(.Pos.Map).SinInviOcul)
                
146                 Case Else
148                     Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'SINMAGIA', 'SININVI', 'NOPKS', 'NOCIUD'", e_FontTypeNames.FONTTYPE_INFO)

                End Select

            End If

        End With

        Exit Sub

errHandler:
150     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoRestricted", Erl)
152

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
    
100     With UserList(UserIndex)
  
102         nomagic = Reader.ReadBool
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            End If

        End With
        
        Exit Sub

HandleChangeMapInfoNoMagic_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoMagic", Erl)
110
        
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
    
100     With UserList(UserIndex)

102         noinvi = Reader.ReadBool()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            End If

        End With
        
        Exit Sub

HandleChangeMapInfoNoInvi_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoInvi", Erl)
110
        
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChangeMapInfoNoResu_Err

        Dim noresu As Boolean
    
100     With UserList(UserIndex)

102         noresu = Reader.ReadBool()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            End If

        End With
        
        Exit Sub

HandleChangeMapInfoNoResu_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoResu", Erl)
110
        
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

        On Error GoTo errHandler

        Dim tStr As String
    
100     With UserList(UserIndex)

102         tStr = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
        
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
            
108                 Call LogGM(.Name, .Name & " ha cambiado la informacion del Terreno del mapa.")
                
110                 MapInfo(UserList(UserIndex).Pos.Map).terrain = tStr
                
112                 Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).terrain, e_FontTypeNames.FONTTYPE_INFO)
                
                Else
            
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", e_FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
122

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
    
        On Error GoTo errHandler

        Dim tStr As String
    
100     With UserList(UserIndex)

102         tStr = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
        
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
            
108                 Call LogGM(.Name, .Name & " ha cambiado la informacion de la Zona del mapa.")
110                 MapInfo(UserList(UserIndex).Pos.Map).zone = tStr
112                 Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).zone, e_FontTypeNames.FONTTYPE_INFO)
                
                Else
            
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", e_FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'DUNGEON' ya que al ingresarlo, NO se sentirí el efecto de la lluvia en este mapa.", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With

        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
122

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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
106         Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
108         Call WriteConsoleMsg(UserIndex, "Mapa Guardado", e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleSaveMap_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSaveMap", Erl)
112
        
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
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim guild As String
102             guild = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call modGuilds.GMEscuchaClan(UserIndex, guild)

            End If

        End With

        Exit Sub

errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
110

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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub

104         Call LogGM(.Name, .Name & " ha hecho un backup")
        
106         Call ES.DoBackUp 'Sino lo confunde con la id del paquete

        End With
        
        Exit Sub

HandleDoBackUp_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
110
        
End Sub

''
' Handle the "AlterName" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: WyroX  -  11/07/2021
    'Change user name
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)

        'Reads the userName and newUser Packets
        Dim UserName     As String
        Dim NewName      As String
        Dim TargetUI     As Integer
        Dim GuildIndex   As Integer

        UserName = UCase$(Reader.ReadString8())
        NewName = Reader.ReadString8()

        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub

        If LenB(UserName) = 0 Or LenB(NewName) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        TargetUI = NameIndex(UserName)
        
        If TargetUI > 0 Then
            If UserList(TargetUI).GuildIndex > 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            If Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " es inexistente.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            GuildIndex = GetUserGuildIndexDatabase(UserName)
            
            If GuildIndex > 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If PersonajeExiste(NewName) Then
            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call ChangeNameDatabase(UserName, NewName)

        Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", e_FontTypeNames.FONTTYPE_INFO)

        Call SavePenaDatabase(UserName, .Name & ": nombre cambiado de """ & UserName & """ a """ & NewName & """. " & Date & " " & Time)
        Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " cambió el nombre del usuario """ & UserName & """ por """ & NewName & """.", e_FontTypeNames.FONTTYPE_GM))
        Call LogGM(.Name, "Ha cambiado de nombre al usuario """ & UserName & """. Ahora se llama """ & NewName & """.")
        
        If TargetUI > 0 Then
            UserList(TargetUI).Name = NewName
            Call RefreshCharStatus(TargetUI)
        End If

    End With

    Exit Sub

errHandler:
150     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAlterName", Erl)
152

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

100     With UserList(UserIndex)

            Dim NpcIndex As Integer
102         NpcIndex = Reader.ReadInt16()

            If Not EsGM(UserIndex) Then Exit Sub
        
104         If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Nos fijamos si es pretoriano.
108         If NpcList(NpcIndex).NPCtype = e_NPCType.Pretoriano Then
110             Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearPretoianos MAPA X Y.", e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
112         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
114         If NpcIndex <> 0 Then
116             Call LogGM(.Name, "Sumoneo a " & NpcList(NpcIndex).Name & " en mapa " & .Pos.Map)

            End If

        End With
        
        Exit Sub

HandleCreateNPC_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPC", Erl)
120
        
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

100     With UserList(UserIndex)

            Dim NpcIndex As Integer
        
102         NpcIndex = Reader.ReadInt16()

            If Not EsGM(UserIndex) Then Exit Sub
        
104         If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
110         If NpcIndex <> 0 Then
112             Call LogGM(.Name, "Sumoneo con respawn " & NpcList(NpcIndex).Name & " en mapa " & .Pos.Map)

            End If

        End With
        
        Exit Sub

HandleCreateNPCWithRespawn_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPCWithRespawn", Erl)
116
        
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

100     With UserList(UserIndex)
        
        
        
            Dim Index    As Byte
            Dim ObjIndex As Integer
        
102         Index = Reader.ReadInt8()
104         ObjIndex = Reader.ReadInt16()
        
106         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
108         Select Case Index

                Case 1
                    ' ArmaduraImperial1 = objindex
            
110             Case 2
                    ' ArmaduraImperial2 = objindex
            
112             Case 3
                    ' ArmaduraImperial3 = objindex
            
114             Case 4

                    ' TunicaMagoImperial = objindex
            End Select

        End With
        
        Exit Sub

HandleImperialArmour_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleImperialArmour", Erl)
118
        
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

100     With UserList(UserIndex)

            Dim Index    As Byte
            Dim ObjIndex As Integer
        
102         Index = Reader.ReadInt8()
104         ObjIndex = Reader.ReadInt16()
        
106         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
108         Select Case Index

                Case 1
                    '   ArmaduraCaos1 = objindex
            
110             Case 2
                    '   ArmaduraCaos2 = objindex
            
112             Case 3
                    '   ArmaduraCaos3 = objindex
            
114             Case 4

                    '  TunicaMagoCaos = objindex
            End Select

        End With
        
        Exit Sub

HandleChaosArmour_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosArmour", Erl)
118
        
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
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         If .flags.Navegando = 1 Then
108             .flags.Navegando = 0
            
            Else
110             .flags.Navegando = 1

            End If
        
            'Tell the client that we are navigating.
112         Call WriteNavigateToggle(UserIndex)

        End With
        
        Exit Sub

HandleNavigateToggle_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
116
        
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
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         If ServerSoloGMs > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", e_FontTypeNames.FONTTYPE_INFO)
108             ServerSoloGMs = 0
            
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", e_FontTypeNames.FONTTYPE_INFO)
112             ServerSoloGMs = 1

            End If

        End With
        
        Exit Sub

HandleServerOpenToUsersToggle_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerOpenToUsersToggle", Erl)
116
        
End Sub

''
' Handle the "Participar" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleParticipar(ByVal UserIndex As Integer)
        On Error GoTo HandleParticipar_Err

        Dim handle As Integer
    
100     With UserList(UserIndex)

102         If Torneo.HayTorneoaActivo = False Then
104             Call WriteConsoleMsg(UserIndex, "No hay ningún evento disponible.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                   
106         If .flags.EnTorneo Then
108             Call WriteConsoleMsg(UserIndex, "Ya estás participando.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
110         If .Stats.ELV > Torneo.nivelmaximo Then
112             Call WriteConsoleMsg(UserIndex, "El nivel máximo para participar es " & Torneo.nivelmaximo & ".", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
114         If .Stats.ELV < Torneo.NivelMinimo Then
116             Call WriteConsoleMsg(UserIndex, "El nivel mínimo para participar es " & Torneo.NivelMinimo & ".", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
118         If .Stats.GLD < Torneo.costo Then
120             Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro para ingresar.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
122         If .clase = Mage And Torneo.mago = 0 Then
124             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
126         If .clase = Cleric And Torneo.clerico = 0 Then
128             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
130         If .clase = Warrior And Torneo.guerrero = 0 Then
132             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
134         If .clase = Bard And Torneo.bardo = 0 Then
136             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
138         If .clase = Assasin And Torneo.asesino = 0 Then
140             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
142         If .clase = Druid And Torneo.druido = 0 Then
144             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
146         If .clase = Paladin And Torneo.Paladin = 0 Then
148             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
150         If .clase = Hunter And Torneo.cazador = 0 Then
152             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
154         If .clase = Trabajador And Torneo.Trabajador = 0 Then
156             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
158         If .clase = e_Class.Thief And Torneo.Ladron = 0 Then
160             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
162         If .clase = e_Class.Bandit And Torneo.Bandido = 0 Then
164             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
166         If .clase = e_Class.Pirat And Torneo.Pirata = 0 Then
168             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
170         If Torneo.Participantes = Torneo.cupos Then
172             Call WriteConsoleMsg(UserIndex, "Los cupos ya estan llenos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
  
174         Call ParticiparTorneo(UserIndex)

        End With
        
        Exit Sub

HandleParticipar_Err:
176     Call TraceError(Err.Number, Err.Description, "Protocol.HandleParticipar", Erl)
178
        
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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "/CONDEN " & UserName)
            
108             tUser = NameIndex(UserName)

110             If tUser > 0 Then Call VolverCriminal(tUser)

            End If

        End With

        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTurnCriminal", Erl)
114

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

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "/RAJAR " & UserName)
            
108             tUser = NameIndex(UserName)
            
110             If tUser > 0 Then Call ResetFacciones(tUser)

            End If

        End With

        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetFactions", Erl)
114

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

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim UserName   As String
            Dim GuildIndex As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
108             GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
110             If GuildIndex = 0 Then
112                 Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", e_FontTypeNames.FONTTYPE_INFO)
                Else
114                 Call WriteConsoleMsg(UserIndex, "Expulsado.", e_FontTypeNames.FONTTYPE_INFO)
116                 Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", e_FontTypeNames.FONTTYPE_GUILD))

                End If

            End If

        End With

        Exit Sub

errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCharFromGuild", Erl)
120

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "Mensaje de sistema:" & Message)
            
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))

            End If

        End With

        Exit Sub

errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSystemMessage", Erl)
112

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

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim newMOTD           As String

            Dim auxiliaryString() As String

            Dim LoopC             As Long
        
102         newMOTD = Reader.ReadString8()
104         auxiliaryString = Split(newMOTD, vbCrLf)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
108             Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
110             MaxLines = UBound(auxiliaryString()) + 1
            
112             ReDim MOTD(1 To MaxLines)
            
114             Call WriteVar(DatPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
116             For LoopC = 1 To MaxLines
118                 Call WriteVar(DatPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
120                 MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
122             Next LoopC
            
124             Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        Exit Sub

errHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetMOTD", Erl)
128

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
100     With UserList(UserIndex)
  
102         If (.flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

            Dim auxiliaryString As String

            Dim LoopC           As Long
        
104         For LoopC = LBound(MOTD()) To UBound(MOTD())
106             auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
108         Next LoopC
        
110         If Len(auxiliaryString) >= 2 Then
        
112             If Right$(auxiliaryString, 2) = vbCrLf Then
114                 auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
                End If

            End If
        
116         Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)

        End With
        
        Exit Sub

HandleChangeMOTD_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMOTD", Erl)
120
        
End Sub

''
' Handle the "Ping" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
        
        On Error GoTo HandlePing_Err

            Dim Time As Long
        
102         Time = Reader.ReadInt32()
        
104         Call WritePong(UserIndex, Time + modNetwork.GetTimeOfNextFlush()) ' Correct the time
        Exit Sub

HandlePing_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePing", Erl)
108
        
End Sub

Private Sub HandleQuestionGM(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Consulta       As String
            Dim TipoDeConsulta As String

102         Consulta = Reader.ReadString8()
104         TipoDeConsulta = Reader.ReadString8()

112         Call Ayuda.Push(.Name, Consulta, TipoDeConsulta)
114         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).Name & ".", e_FontTypeNames.FONTTYPE_SERVER))

116         Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", e_FontTypeNames.FONTTYPE_INFOIAO)
        
118         Call LogConsulta(.Name & " (" & TipoDeConsulta & ") " & Consulta)

        End With
    
        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
122

End Sub

Private Sub HandleOfertaInicial(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOfertaInicial_Err
    
100     With UserList(UserIndex)

            Dim Oferta As Long
102             Oferta = Reader.ReadInt32()
        
104         If UserList(UserIndex).flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                
                Exit Sub

            End If

108         If .flags.TargetNPC < 1 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

112         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Subastador Then
114             Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
116         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 2 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del subastador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
120         If .flags.Subastando = False Then
122             Call WriteChatOverHead(UserIndex, "Oye amigo, tu no podés decirme cual es la oferta inicial.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
124         If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
126             Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
128         If .flags.Subastando = True Then
130             UserList(UserIndex).Counters.TiempoParaSubastar = 0
132             Subasta.OfertaInicial = Oferta
134             Subasta.MejorOferta = 0
136             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " está subastando: " & ObjData(Subasta.ObjSubastado).Name & " (Cantidad: " & Subasta.ObjSubastadoCantidad & " ) - con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas. Escribe /OFERTAR (cantidad) para participar.", e_FontTypeNames.FONTTYPE_SUBASTA))
138             .flags.Subastando = False
140             Subasta.HaySubastaActiva = True
142             Subasta.Subastador = .Name
144             Subasta.MinutosDeSubasta = 5
146             Subasta.TiempoRestanteSubasta = 300
148             Call LogearEventoDeSubasta("#################################################################################################################################################################################################")
150             Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
152             Call LogearEventoDeSubasta(.Name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & " y con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas.")
154             frmMain.SubastaTimer.Enabled = True
156             Call WarpUserChar(UserIndex, 14, 27, 64, True)

                'lalala toda la bola de los timerrr
            End If

        End With
        
        Exit Sub

HandleOfertaInicial_Err:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOfertaInicial", Erl)
160
        
End Sub

Private Sub HandleOfertaDeSubasta(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Oferta   As Long
            Dim ExOferta As Long
        
102         Oferta = Reader.ReadInt32()
        
104         If Subasta.HaySubastaActiva = False Then
106             Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", e_FontTypeNames.FONTTYPE_INFOIAO)
            
                Exit Sub

            End If
               
108         If Oferta < Subasta.MejorOferta + 100 Then
110             Call WriteConsoleMsg(UserIndex, "Debe haber almenos una diferencia de 100 monedas a la ultima oferta!", e_FontTypeNames.FONTTYPE_INFOIAO)
            
                Exit Sub

            End If
        
112         If .Name = Subasta.Subastador Then
114             Call WriteConsoleMsg(UserIndex, "No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...", e_FontTypeNames.FONTTYPE_INFOIAO)
            
                Exit Sub

            End If
        
116         If .Stats.GLD >= Oferta Then

                'revisar que pasa si el usuario que oferto antes esta offline
                'Devolvemos el oro al usuario que oferto antes...(si es que hubo oferta)
118             If Subasta.HuboOferta = True Then
120                 ExOferta = NameIndex(Subasta.Comprador)
122                 UserList(ExOferta).Stats.GLD = UserList(ExOferta).Stats.GLD + Subasta.MejorOferta
124                 Call WriteUpdateGold(ExOferta)

                End If
            
126             Subasta.MejorOferta = Oferta
128             Subasta.Comprador = .Name
            
130             .Stats.GLD = .Stats.GLD - Oferta
132             Call WriteUpdateGold(UserIndex)
            
134             If Subasta.TiempoRestanteSubasta < 60 Then
136                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .Name & " (Ofrece " & PonerPuntos(Oferta) & " monedas de oro) - Tiempo Extendido. Escribe /SUBASTA para mas información.", e_FontTypeNames.FONTTYPE_SUBASTA))
138                 Call LogearEventoDeSubasta(.Name & ": Mejoro la oferta en el ultimo minuto ofreciendo " & PonerPuntos(Oferta) & " monedas.")
140                 Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta + 30
                Else
142                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .Name & " (Ofrece " & PonerPuntos(Oferta) & " monedas de oro). Escribe /SUBASTA para mas información.", e_FontTypeNames.FONTTYPE_SUBASTA))
144                 Call LogearEventoDeSubasta(.Name & ": Mejoro la oferta ofreciendo " & PonerPuntos(Oferta) & " monedas.")
146                 Subasta.HuboOferta = True
148                 Subasta.PosibleCancelo = False

                End If

            Else
150             Call WriteConsoleMsg(UserIndex, "No posees esa cantidad de oro.", e_FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End With
    
        Exit Sub

errHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
154

End Sub

Private Sub HandleSetSpeed(ByVal UserIndex As Integer)
        
        Dim Speed As Single
        
        On Error GoTo HandleGlobalOnOff_Err
        
        Speed = Reader.ReadReal32()
        
        'Author: Pablo Mercavides
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then Exit Sub
            
            UserList(UserIndex).Char.speeding = Speed
            
            Call WriteVelocidadToggle(Speed)
        
        End With
        
        Exit Sub

HandleGlobalOnOff_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
118
        
End Sub

Private Sub HandleGlobalMessage(ByVal UserIndex As Integer)
        
        Dim TActual     As Long
        Dim ElapsedTime As Long

100     TActual = GetTickCount()
102     ElapsedTime = TActual - UserList(UserIndex).Counters.MensajeGlobal
                
        On Error GoTo errHandler

104     With UserList(UserIndex)

            Dim chat As String

106         chat = Reader.ReadString8()

108         If .flags.Silenciado = 1 Then
110             Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
                'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", e_FontTypeNames.FONTTYPE_VENENO)
        
112         ElseIf ElapsedTime < IntervaloMensajeGlobal Then
114             Call WriteConsoleMsg(UserIndex, "No puedes escribir mensajes globales tan rápido.", e_FontTypeNames.FONTTYPE_WARNING)
        
            Else
116             UserList(UserIndex).Counters.MensajeGlobal = TActual

118             If EstadoGlobal Then
120                 If LenB(chat) <> 0 Then
                        'Analize chat...
122                     Call Statistics.ParseChat(chat)

                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Integer

124                     For i = 1 To UBound(.flags.ChatHistory) - 1
126                         .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
128                     .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat

130                     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[" & .Name & "] " & chat, e_FontTypeNames.FONTTYPE_GLOBAL))

                        'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                        'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbBlue & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                    End If

                Else
132                 Call WriteConsoleMsg(UserIndex, "El global se encuentra Desactivado.", e_FontTypeNames.FONTTYPE_GLOBAL)

                End If

            End If
    
        End With
    
        Exit Sub

errHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
136

End Sub

Public Sub HandleGlobalOnOff(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGlobalOnOff_Err

        'Author: Pablo Mercavides
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub
        
104         Call LogGM(.Name, " activo al Chat Global a las " & Now)
        
106         If EstadoGlobal = False Then
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Chat general habilitado. Escribe" & Chr(34) & "/CONSOLA" & Chr(34) & " o " & Chr(34) & ";" & Chr(34) & " y su mensaje para utilizarlo.", e_FontTypeNames.FONTTYPE_SERVER))
110             EstadoGlobal = True
            Else
112             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Chat General deshabilitado.", e_FontTypeNames.FONTTYPE_SERVER))
114             EstadoGlobal = False

            End If
        
        End With
        
        Exit Sub

HandleGlobalOnOff_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
118
        
End Sub

Private Sub HandleIngresarConCuenta(ByVal UserIndex As Integer)

        On Error GoTo errHandler

        Dim Version As String
    
100     With UserList(UserIndex)

            Dim CuentaEmail    As String
            Dim CuentaPassword As String
            Dim MD5            As String
        
102         CuentaEmail = Reader.ReadString8()
104         CuentaPassword = Reader.ReadString8()
106         Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
112         MD5 = Reader.ReadString8()
        
            #If DEBUGGING = False Then
    
114             If Not VersionOK(Version) Then
116                 Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
118                 Call CloseSocket(UserIndex)
                    Exit Sub
        
                End If
    
            #End If
    
120         Call EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MD5)

        End With
    
        Exit Sub

errHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIngresarConCuenta", Erl)
150

End Sub

Private Sub HandleBorrarPJ(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)

            Dim UserDelete     As String
            Dim CuentaEmail    As String
            Dim CuentaPassword As String
            Dim MD5            As String
            Dim Version        As String
        
102         UserDelete = Reader.ReadString8()
104         CuentaEmail = Reader.ReadString8()
106         CuentaPassword = Reader.ReadString8()
108         Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
114         MD5 = Reader.ReadString8()
        
            #If DEBUGGING = False Then
116             If Not VersionOK(Version) Then
118                 Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
120                 Call CloseSocket(UserIndex)
                    Exit Sub
                End If
            #End If
        
122         If Not EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MD5) Then
124             Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Dim RS As Recordset
            Set RS = Query("SELECT account_id, level, is_banned from user WHERE name = ?", UserDelete)
            
            If (RS Is Nothing) Then
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            If (RS!account_id <> UserList(UserIndex).AccountID) Then
128             Call LogHackAttemp(CuentaEmail & "[" & UserList(UserIndex).IP & "] intentó borrar el pj " & UserDelete)
130             Call CloseSocket(UserIndex)
                Exit Sub
            End If

132         If (RS!level >= 25) Then
134             Call WriteShowMessageBox(UserIndex, "No puedes eliminar un personaje mayor a nivel 25.")
                Exit Sub
            End If

136         If (RS!is_banned) Then
138             Call WriteShowMessageBox(UserIndex, "No puedes eliminar un personaje baneado.")
                Exit Sub
            End If

            'HarThaoS: Si teine clan y es leader no lo puedo eliminar
140         If PersonajeEsLeader(UserDelete) Then
142             Call WriteShowMessageBox(UserIndex, "No puedes eliminar el personaje por ser líder de un clan.")
                Exit Sub
            End If

            ' Si está online el personaje a borrar, lo kickeo para prevenir dupeos.
            Dim targetUserIndex As Integer
144         targetUserIndex = NameIndex(UserDelete)

            'HarThaoS: Me fijo si tiene clan y me traigo el nombre del clan

146         If targetUserIndex > 0 Then
148             Call LogHackAttemp("Se trató de eliminar al personaje " & UserDelete & " cuando este estaba conectado desde la IP " & UserList(UserIndex).IP)
150             Call CloseSocket(targetUserIndex)

            End If

152         Call BorrarUsuarioDatabase(UserDelete)
154         Call WritePersonajesDeCuenta(UserIndex, Nothing)
  
        End With
    
        Exit Sub

errHandler:
156     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
158

End Sub

Private Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim Seconds As Byte
        
102         Seconds = Reader.ReadInt8()

104         If Not .flags.Privilegios And e_PlayerType.user Then
106             CuentaRegresivaTimer = Seconds
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Empezando cuenta regresiva desde: " & Seconds & " segundos...!", e_FontTypeNames.FONTTYPE_GUILD))
        
            
            End If

        End With
    
        Exit Sub

errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaRegresiva", Erl)
112

End Sub

Private Sub HandlePossUser(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
        
102         UserName = Reader.ReadString8()

104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) = 0 Then
106             If NameIndex(UserName) <= 0 Then

110                     If Not SetPositionDatabase(UserName, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then
112                         Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no existe.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

116                 Call WriteConsoleMsg(UserIndex, "Servidor » Acción realizada con exito! La nueva posicion de " & UserName & " es: " & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.Y & "...", e_FontTypeNames.FONTTYPE_INFO)

                Else
118                 Call WriteConsoleMsg(UserIndex, "Servidor » El usuario debe estar deslogueado para dicha solicitud!", e_FontTypeNames.FONTTYPE_INFO)

                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
    
        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePossUser", Erl)
124

End Sub

Private Sub HandleDuel(ByVal UserIndex As Integer)
    
        On Error GoTo errHandler
        
        Dim Players         As String
        Dim Bet             As Long
        Dim PocionesMaximas As Integer
        Dim CaenItems       As Boolean

100     With UserList(UserIndex)

102         Players = Reader.ReadString8
104         Bet = Reader.ReadInt32
106         PocionesMaximas = Reader.ReadInt16
108         CaenItems = Reader.ReadBool

110         Call CrearReto(UserIndex, Players, Bet, PocionesMaximas, CaenItems)

        End With
    
        Exit Sub
    
errHandler:

112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDuel", Erl)
114

End Sub

Private Sub HandleAcceptDuel(ByVal UserIndex As Integer)

        On Error GoTo errHandler
        
        Dim Offerer As String

100     With UserList(UserIndex)

102         Offerer = Reader.ReadString8

104         Call AceptarReto(UserIndex, Offerer)

        End With
    
        Exit Sub
    
errHandler:

106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptDuel", Erl)
108

End Sub

Private Sub HandleCancelDuel(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         Reader.ReadInt16

104         If .flags.SolicitudReto.estado <> e_SolicitudRetoEstado.Libre Then
106             Call CancelarSolicitudReto(UserIndex, .Name & " ha cancelado la solicitud.")

108         ElseIf .flags.AceptoReto > 0 Then
110             Call CancelarSolicitudReto(.flags.AceptoReto, .Name & " ha cancelado su admisión.")

            End If

        End With

End Sub

Private Sub HandleQuitDuel(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .flags.EnReto Then
104             Call AbandonarReto(UserIndex)
            End If

        End With

End Sub

Private Sub HandleNieveToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleNieveToggle_Err

        'Author: Pablo Mercavides
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         Call LogGM(.Name, "/NIEVE")
        
108         Nebando = Not Nebando
        
110         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

        End With
        
        Exit Sub

HandleNieveToggle_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
114
        
End Sub

Private Sub HandleNieblaToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleNieblaToggle_Err

        'Author: Pablo Mercavides
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         Call LogGM(.Name, "/NIEBLA")
108         Call ResetMeteo

        End With
        
        Exit Sub

HandleNieblaToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
112
        
End Sub

Private Sub HandleTransFerGold(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Cantidad As Long
            Dim tUser    As Integer
        
102         Cantidad = Reader.ReadInt32()
104         UserName = Reader.ReadString8()

            ' WyroX: Chequeos de seguridad... Estos chequeos ya se hacen en el cliente, pero si no se hacen se puede duplicar oro...

            ' Cantidad válida?
106         If Cantidad <= 0 Then Exit Sub

            ' Tiene el oro?
108         If .Stats.Banco < Cantidad Then Exit Sub
            
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

118         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then Exit Sub
            
120         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
122             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

124         tUser = NameIndex(UserName)

            ' Enviar a vos mismo?
126         If tUser = UserIndex Then
128             Call WriteChatOverHead(UserIndex, "¡No puedo enviarte oro a vos mismo!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
    
130         If Not EsGM(UserIndex) Then

132             If tUser <= 0 Then

136                     If Not AddOroBancoDatabase(UserName, Cantidad) Then
138                         Call WriteChatOverHead(UserIndex, "El usuario no existe.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                            Exit Sub

                        End If


                Else
148                 UserList(tUser).Stats.Banco = UserList(tUser).Stats.Banco + val(Cantidad) 'Se lo damos al otro.

                End If
                
150             UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
    
152             Call WriteChatOverHead(UserIndex, "¡El envío se ha realizado con éxito! Gracias por utilizar los servicios de Finanzas Goliath", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
'154             Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave("173", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
            Else
156             Call WriteChatOverHead(UserIndex, "Los administradores no pueden transferir oro.", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
158             Call LogGM(.Name, "Quizo transferirle oro a: " & UserName)
            
            End If

        End With
    
        Exit Sub

errHandler:
160     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
162

End Sub

Private Sub HandleMoveItem(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim SlotViejo As Byte
            Dim SlotNuevo As Byte
        
102         SlotViejo = Reader.ReadInt8()
104         SlotNuevo = Reader.ReadInt8()
        
            Dim Objeto    As t_Obj
            Dim Equipado  As Boolean
            Dim Equipado2 As Boolean
            Dim Equipado3 As Boolean
        
106         If (SlotViejo > .CurrentInventorySlots) Or (SlotNuevo > .CurrentInventorySlots) Then
108             Call WriteConsoleMsg(UserIndex, "Espacio no desbloqueado.", e_FontTypeNames.FONTTYPE_INFOIAO)
            
            Else
    
110             If .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex Then
112                 .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotNuevo).amount + .Invent.Object(SlotViejo).amount
                    
                    Dim Excedente As Integer
114                 Excedente = .Invent.Object(SlotNuevo).amount - MAX_INVENTORY_OBJS

116                 If Excedente > 0 Then
118                     .Invent.Object(SlotViejo).amount = Excedente
120                     .Invent.Object(SlotNuevo).amount = MAX_INVENTORY_OBJS
                    Else

122                     If .Invent.Object(SlotViejo).Equipped = 1 Then
124                         .Invent.Object(SlotNuevo).Equipped = 1

                        End If
                    
126                     .Invent.Object(SlotViejo).ObjIndex = 0
128                     .Invent.Object(SlotViejo).amount = 0
130                     .Invent.Object(SlotViejo).Equipped = 0
                    
                        'Cambiamos si alguno es un anillo
132                     If .Invent.DañoMagicoEqpSlot = SlotViejo Then
134                         .Invent.DañoMagicoEqpSlot = SlotNuevo

                        End If

136                     If .Invent.ResistenciaEqpSlot = SlotViejo Then
138                         .Invent.ResistenciaEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un armor
140                     If .Invent.ArmourEqpSlot = SlotViejo Then
142                         .Invent.ArmourEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un barco
144                     If .Invent.BarcoSlot = SlotViejo Then
146                         .Invent.BarcoSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es una montura
148                     If .Invent.MonturaSlot = SlotViejo Then
150                         .Invent.MonturaSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un casco
152                     If .Invent.CascoEqpSlot = SlotViejo Then
154                         .Invent.CascoEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un escudo
156                     If .Invent.EscudoEqpSlot = SlotViejo Then
158                         .Invent.EscudoEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es munición
160                     If .Invent.MunicionEqpSlot = SlotViejo Then
162                         .Invent.MunicionEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un arma
164                     If .Invent.WeaponEqpSlot = SlotViejo Then
166                         .Invent.WeaponEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un nudillo
168                     If .Invent.NudilloSlot = SlotViejo Then
170                         .Invent.NudilloSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un magico
172                     If .Invent.MagicoSlot = SlotViejo Then
174                         .Invent.MagicoSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es una herramienta
176                     If .Invent.HerramientaEqpSlot = SlotViejo Then
178                         .Invent.HerramientaEqpSlot = SlotNuevo

                        End If

                    End If
                
                Else

180                 If .Invent.Object(SlotNuevo).ObjIndex <> 0 Then
182                     Objeto.amount = .Invent.Object(SlotViejo).amount
184                     Objeto.ObjIndex = .Invent.Object(SlotViejo).ObjIndex
                    
186                     If .Invent.Object(SlotViejo).Equipped = 1 Then
188                         Equipado = True
    
                        End If
                    
190                     If .Invent.Object(SlotNuevo).Equipped = 1 Then
192                         Equipado2 = True
    
                        End If
                    
                        '  If .Invent.Object(SlotNuevo).Equipped = 1 And .Invent.Object(SlotViejo).Equipped = 1 Then
                        '     Equipado3 = True
                        ' End If
                    
194                     .Invent.Object(SlotViejo).ObjIndex = .Invent.Object(SlotNuevo).ObjIndex
196                     .Invent.Object(SlotViejo).amount = .Invent.Object(SlotNuevo).amount
                    
198                     .Invent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
200                     .Invent.Object(SlotNuevo).amount = Objeto.amount
                    
202                     If Equipado Then
204                         .Invent.Object(SlotNuevo).Equipped = 1
                        Else
206                         .Invent.Object(SlotNuevo).Equipped = 0
    
                        End If
                                    
208                     If Equipado2 Then
210                         .Invent.Object(SlotViejo).Equipped = 1
                        Else
212                         .Invent.Object(SlotViejo).Equipped = 0
    
                        End If
    
                    End If
    
                    'Cambiamos si alguno es un anillo
214                 If .Invent.DañoMagicoEqpSlot = SlotViejo Then
216                     .Invent.DañoMagicoEqpSlot = SlotNuevo
218                 ElseIf .Invent.DañoMagicoEqpSlot = SlotNuevo Then
220                     .Invent.DañoMagicoEqpSlot = SlotViejo

                    End If

222                 If .Invent.ResistenciaEqpSlot = SlotViejo Then
224                     .Invent.ResistenciaEqpSlot = SlotNuevo
226                 ElseIf .Invent.ResistenciaEqpSlot = SlotNuevo Then
228                     .Invent.ResistenciaEqpSlot = SlotViejo

                    End If
                    
                    'Cambiamos si alguno es un armor
230                 If .Invent.ArmourEqpSlot = SlotViejo Then
232                     .Invent.ArmourEqpSlot = SlotNuevo
234                 ElseIf .Invent.ArmourEqpSlot = SlotNuevo Then
236                     .Invent.ArmourEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un barco
238                 If .Invent.BarcoSlot = SlotViejo Then
240                     .Invent.BarcoSlot = SlotNuevo
242                 ElseIf .Invent.BarcoSlot = SlotNuevo Then
244                     .Invent.BarcoSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es una montura
246                 If .Invent.MonturaSlot = SlotViejo Then
248                     .Invent.MonturaSlot = SlotNuevo
250                 ElseIf .Invent.MonturaSlot = SlotNuevo Then
252                     .Invent.MonturaSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un casco
254                 If .Invent.CascoEqpSlot = SlotViejo Then
256                     .Invent.CascoEqpSlot = SlotNuevo
258                 ElseIf .Invent.CascoEqpSlot = SlotNuevo Then
260                     .Invent.CascoEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un escudo
262                 If .Invent.EscudoEqpSlot = SlotViejo Then
264                     .Invent.EscudoEqpSlot = SlotNuevo
266                 ElseIf .Invent.EscudoEqpSlot = SlotNuevo Then
268                     .Invent.EscudoEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es munición
270                 If .Invent.MunicionEqpSlot = SlotViejo Then
272                     .Invent.MunicionEqpSlot = SlotNuevo
274                 ElseIf .Invent.MunicionEqpSlot = SlotNuevo Then
276                     .Invent.MunicionEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un arma
278                 If .Invent.WeaponEqpSlot = SlotViejo Then
280                     .Invent.WeaponEqpSlot = SlotNuevo
282                 ElseIf .Invent.WeaponEqpSlot = SlotNuevo Then
284                     .Invent.WeaponEqpSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es un nudillo
286                 If .Invent.NudilloSlot = SlotViejo Then
288                     .Invent.NudilloSlot = SlotNuevo
290                 ElseIf .Invent.NudilloSlot = SlotNuevo Then
292                     .Invent.NudilloSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es un magico
294                 If .Invent.MagicoSlot = SlotViejo Then
296                     .Invent.MagicoSlot = SlotNuevo
298                 ElseIf .Invent.MagicoSlot = SlotNuevo Then
300                     .Invent.MagicoSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es una herramienta
302                 If .Invent.HerramientaEqpSlot = SlotViejo Then
304                     .Invent.HerramientaEqpSlot = SlotNuevo
306                 ElseIf .Invent.HerramientaEqpSlot = SlotNuevo Then
308                     .Invent.HerramientaEqpSlot = SlotViejo
    
                    End If
                
310                 If Objeto.ObjIndex = 0 Then
312                     .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex
314                     .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotViejo).amount
316                     .Invent.Object(SlotNuevo).Equipped = .Invent.Object(SlotViejo).Equipped
                            
318                     .Invent.Object(SlotViejo).ObjIndex = 0
320                     .Invent.Object(SlotViejo).amount = 0
322                     .Invent.Object(SlotViejo).Equipped = 0
    
                    End If
                    
                End If
                
324             Call UpdateUserInv(False, UserIndex, SlotViejo)
326             Call UpdateUserInv(False, UserIndex, SlotNuevo)

            End If

        End With
    
        Exit Sub

errHandler:
328     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveItem", Erl)
330

End Sub

Private Sub HandleBovedaMoveItem(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim SlotViejo As Byte
            Dim SlotNuevo As Byte
        
102         SlotViejo = Reader.ReadInt8()
104         SlotNuevo = Reader.ReadInt8()
        
            Dim Objeto    As t_Obj
            Dim Equipado  As Boolean
            Dim Equipado2 As Boolean
            Dim Equipado3 As Boolean
        
106         Objeto.ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex
108         Objeto.amount = UserList(UserIndex).BancoInvent.Object(SlotViejo).amount
        
110         UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex
112         UserList(UserIndex).BancoInvent.Object(SlotViejo).amount = UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount
         
114         UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
116         UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount = Objeto.amount
    
            'Actualizamos el banco
118         Call UpdateBanUserInv(False, UserIndex, SlotViejo)
120         Call UpdateBanUserInv(False, UserIndex, SlotNuevo)

        End With
    
        Exit Sub
    
        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBovedaMoveItem", Erl)
124

End Sub

Private Sub HandleQuieroFundarClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

102         If UserList(UserIndex).flags.Privilegios And e_PlayerType.Consejero Then Exit Sub

104         If UserList(UserIndex).GuildIndex > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Ya perteneces a un clan, no podés fundar otro.", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

108         If UserList(UserIndex).Stats.ELV < 35 Or UserList(UserIndex).Stats.UserSkills(e_Skill.liderazgo) < 100 Then
110             Call WriteConsoleMsg(UserIndex, "Para fundar un clan debes ser nivel 35, tener 100 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1).", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

112         If Not TieneObjetos(407, 1, UserIndex) Or Not TieneObjetos(408, 1, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "Para fundar un clan debes tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1).", e_FontTypeNames.FONTTYPE_INFOIAO)
                'Exit Sub

            End If

116         Call WriteConsoleMsg(UserIndex, "Servidor » ¡Comenzamos a fundar el clan! Ingresa todos los datos solicitados.", e_FontTypeNames.FONTTYPE_INFOIAO)
        
118         Call WriteShowFundarClanForm(UserIndex)

        End With
    
        Exit Sub
    
        Exit Sub

errHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuieroFundarClan", Erl)
122

End Sub

Private Sub HandleLlamadadeClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim refError   As String
            Dim clan_nivel As Byte

102         If .GuildIndex <> 0 Then
104             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)

106             If clan_nivel >= 2 Then
108                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> [" & .Name & "] solicita apoyo de su clan en " & DarNameMapa(.Pos.Map) & " (" & .Pos.Map & "-" & .Pos.x & "-" & .Pos.Y & "). Puedes ver su ubicación en el mapa del mundo.", e_FontTypeNames.FONTTYPE_GUILD))
110                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
112                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.Pos.Map, .Pos.x, .Pos.Y))

                Else
114                 Call WriteConsoleMsg(UserIndex, "Servidor » El nivel de tu clan debe ser 2 para utilizar esta opción.", e_FontTypeNames.FONTTYPE_INFOIAO)

                End If
            End If

        End With
    
        Exit Sub

errHandler:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLlamadadeClan", Erl)
118

End Sub


Private Sub HandleGenio(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleGenio_Err

100     With UserList(UserIndex)

            'Si no es GM, no pasara nada :P
102         If (.flags.Privilegios And e_PlayerType.user) Then Exit Sub
        
            Dim i As Byte

104         For i = 1 To NUMSKILLS
106             .Stats.UserSkills(i) = 100
108         Next i
        
110         Call WriteConsoleMsg(UserIndex, "Tus skills fueron editados.", e_FontTypeNames.FONTTYPE_INFOIAO)

        End With
        
        Exit Sub

HandleGenio_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGenio", Erl)
        
End Sub

Private Sub HandleCasamiento(ByVal UserIndex As Integer)

        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer

102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
            
106         If .flags.TargetNPC > 0 Then

108             If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Revividor Then
110                 Call WriteConsoleMsg(UserIndex, "Primero haz click sobre un sacerdote.", e_FontTypeNames.FONTTYPE_INFO)

                Else

112                 If Distancia(.Pos, NpcList(.flags.TargetNPC).Pos) > 10 Then
114                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede casarte debido a que estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                        
                    Else
            
116                     If tUser = UserIndex Then
118                         Call WriteConsoleMsg(UserIndex, "No podés casarte contigo mismo.", e_FontTypeNames.FONTTYPE_INFO)
                        
120                     ElseIf .flags.Casado = 1 Then
122                         Call WriteConsoleMsg(UserIndex, "¡Ya estás casado! Debes divorciarte de tu actual pareja para casarte nuevamente.", e_FontTypeNames.FONTTYPE_INFO)
                            
124                     ElseIf UserList(tUser).flags.Casado = 1 Then
126                         Call WriteConsoleMsg(UserIndex, "Tu pareja debe divorciarse antes de tomar tu mano en matrimonio.", e_FontTypeNames.FONTTYPE_INFO)
                            
                        Else

128                         If tUser <= 0 Then
130                             Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", e_FontTypeNames.FONTTYPE_INFO)

                            Else

132                             If UserList(tUser).flags.Candidato = UserIndex Then

134                                 UserList(tUser).flags.Casado = 1
136                                 UserList(tUser).flags.Pareja = UserList(UserIndex).Name
138                                 .flags.Casado = 1
140                                 .flags.Pareja = UserList(tUser).Name

142                                 Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_FXSound.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
144                                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sacerdote de " & DarNameMapa(.Pos.Map) & " celebra el casamiento entre " & UserList(UserIndex).Name & " y " & UserList(tUser).Name & ".", e_FontTypeNames.FONTTYPE_WARNING))
146                                 Call WriteChatOverHead(UserIndex, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
148                                 Call WriteChatOverHead(tUser, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                                
                                Else
                                
150                                 Call WriteChatOverHead(UserIndex, "La solicitud de casamiento a sido enviada a " & UserName & ".", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
152                                 Call WriteConsoleMsg(tUser, .Name & " desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER " & .Name & ".", e_FontTypeNames.FONTTYPE_TALK)

154                                 .flags.Candidato = tUser

                                End If

                            End If

                        End If

                    End If

                End If

            Else
156             Call WriteConsoleMsg(UserIndex, "Primero haz click sobre el sacerdote.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
    
        Exit Sub

errHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCasamiento", Erl)
160

End Sub

Private Sub HandleCrearTorneo(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)
 
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
            Dim Pirata      As Byte
            Dim Ladron      As Byte
            Dim Bandido     As Byte

            Dim Mapa        As Integer
            Dim x           As Byte
            Dim Y           As Byte

            Dim nombre      As String
            Dim reglas      As String

102         NivelMinimo = Reader.ReadInt8
104         nivelmaximo = Reader.ReadInt8
        
106         cupos = Reader.ReadInt8
108         costo = Reader.ReadInt32
        
110         mago = Reader.ReadInt8
112         clerico = Reader.ReadInt8
114         guerrero = Reader.ReadInt8
116         asesino = Reader.ReadInt8
118         bardo = Reader.ReadInt8
120         druido = Reader.ReadInt8
122         Paladin = Reader.ReadInt8
124         cazador = Reader.ReadInt8
126         Trabajador = Reader.ReadInt8
128         Pirata = Reader.ReadInt8
130         Ladron = Reader.ReadInt8
132         Bandido = Reader.ReadInt8

134         Mapa = Reader.ReadInt16
136         x = Reader.ReadInt8
138         Y = Reader.ReadInt8
        
140         nombre = Reader.ReadString8
142         reglas = Reader.ReadString8
  
144         If EsGM(UserIndex) And ((.flags.Privilegios And e_PlayerType.Consejero) = 0) Then
146             Torneo.NivelMinimo = NivelMinimo
148             Torneo.nivelmaximo = nivelmaximo
            
150             Torneo.cupos = cupos
152             Torneo.costo = costo
            
154             Torneo.mago = mago
156             Torneo.clerico = clerico
158             Torneo.guerrero = guerrero
160             Torneo.asesino = asesino
162             Torneo.bardo = bardo
164             Torneo.druido = druido
166             Torneo.Paladin = Paladin
168             Torneo.cazador = cazador
170             Torneo.Trabajador = Trabajador
172             Torneo.Pirata = Pirata
174             Torneo.Ladron = Ladron
176             Torneo.Bandido = Bandido
        
178             Torneo.Mapa = Mapa
180             Torneo.x = x
182             Torneo.Y = Y
            
184             Torneo.nombre = nombre
186             Torneo.reglas = reglas

188             Call IniciarTorneo

            End If

        End With
    
        Exit Sub

errHandler:
190     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCrearTorneo", Erl)
192

End Sub

Private Sub HandleComenzarTorneo(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then

104             Call ComenzarTorneoOk

            End If

        End With
    
        Exit Sub

errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
108

End Sub

Private Sub HandleCancelarTorneo(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then
104             Call ResetearTorneo

            End If

        End With
    
        Exit Sub

errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
108

End Sub

Private Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Tipo As Byte
102             Tipo = Reader.ReadInt8()
  
104         If (.flags.Privilegios And Not (e_PlayerType.Consejero Or e_PlayerType.user)) Then

106             Select Case Tipo

                    Case 0

108                     If Not BusquedaTesoroActiva And BusquedaRegaloActiva = False And BusquedaNpcActiva = False Then
110                         Call PerderTesoro
                        Else

112                         If BusquedaTesoroActiva Then
114                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & "). ¿Quien sera el valiente que lo encuentre?", e_FontTypeNames.FONTTYPE_TALK))
116                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & TesoroNumMapa & "-" & TesoroX & "-" & TesoroY, e_FontTypeNames.FONTTYPE_INFO)
                            Else
118                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

120                 Case 1

122                     If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
124                         Call PerderRegalo
                        Else

126                         If BusquedaRegaloActiva Then
128                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningún valiente fue capaz de encontrar el item misterioso, recuerda que se encuentra en " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & "). ¡Ten cuidado!", e_FontTypeNames.FONTTYPE_TALK))
130                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & RegaloNumMapa & "-" & RegaloX & "-" & RegaloY, e_FontTypeNames.FONTTYPE_INFO)
                            Else
132                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

134                 Case 2

136                     If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                            Dim Pos As t_WorldPos
138                         Pos.Map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
140                         Pos.Y = 50
142                         Pos.x = 50
144                         npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), Pos, True, False, True)
146                         BusquedaNpcActiva = True
                        Else

148                         If BusquedaNpcActiva Then
150                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavía nadie logró matar el NPC que se encuentra en el mapa " & NpcList(npc_index_evento).Pos.Map & ".", e_FontTypeNames.FONTTYPE_TALK))
152                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda de npc activo. El tesoro se encuentra en: " & NpcList(npc_index_evento).Pos.Map & "-" & NpcList(npc_index_evento).Pos.x & "-" & NpcList(npc_index_evento).Pos.Y, e_FontTypeNames.FONTTYPE_INFO)
                            Else
154                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                End Select
            Else
156             Call WriteConsoleMsg(UserIndex, "Servidor » No estas habilitado para hacer Eventos.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

errHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
160

End Sub

Private Sub HandleFlagTrabajar(ByVal UserIndex As Integer)
    
        On Error GoTo errHandler

100     With UserList(UserIndex)

102         .Counters.Trabajando = 0
104         .flags.UsandoMacro = False
106         .flags.TargetObj = 0 ' Sacamos el targer del objeto
108         .flags.UltimoMensaje = 0

        End With
    
        Exit Sub

errHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
112

End Sub

Private Sub HandleEscribiendo(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)
            
            .flags.Escribiendo = Reader.ReadBool()
            
106         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetEscribiendo(.Char.CharIndex, .flags.Escribiendo))

        End With
    
        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
114

End Sub

Private Sub HandleRequestFamiliar(ByVal UserIndex As Integer)
 
        On Error GoTo HandleRequestFamiliar_Err

100     Call WriteFamiliar(UserIndex)
        
        Exit Sub

HandleRequestFamiliar_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestFamiliar", Erl)
104
        
End Sub

Private Sub HandleCompletarAccion(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Accion As Byte
102             Accion = Reader.ReadInt8()
        
104         If .Accion.AccionPendiente = True Then
106             If .Accion.TipoAccion = Accion Then
108                 Call CompletarAccionFin(UserIndex)
                Else
110                 Call WriteConsoleMsg(UserIndex, "Servidor » La acción que solicitas no se corresponde.", e_FontTypeNames.FONTTYPE_SERVER)

                End If

            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Tu no tenias ninguna acción pendiente. ", e_FontTypeNames.FONTTYPE_SERVER)

            End If

        End With
    
        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
116

End Sub

Private Sub HandleInvitarGrupo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", e_FontTypeNames.FONTTYPE_INFO)
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
            
            Else
            
106             If .Grupo.CantidadMiembros <= UBound(.Grupo.Miembros) Then
108                 Call WriteWorkRequestTarget(UserIndex, e_Skill.Grupo)
                Else
110                 Call WriteConsoleMsg(UserIndex, "¡No podés invitar a más personas!", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With
        
        Exit Sub

HandleInvitarGrupo_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvitarGrupo", Erl)
114
    
End Sub

Private Sub HandleMarcaDeClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleMarcaDeClan_Err

100     With UserList(UserIndex)
    
102         If UserList(UserIndex).GuildIndex = 0 Then
                Exit Sub
            End If

104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", e_FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim clan_nivel As Byte

108         clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

110         If clan_nivel < 3 Then
112             Call WriteConsoleMsg(UserIndex, "Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
       
114         Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeClan)
        
        End With
        
        Exit Sub

HandleMarcaDeClan_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeClan", Erl)
118
End Sub

Private Sub HandleMarcaDeGM(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleMarcaDeGM_Err

100     Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeGM)

        Exit Sub

HandleMarcaDeGM_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeGM", Erl)
104
    
End Sub

Private Sub HandleResponderPregunta(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim respuesta As Boolean
            Dim DeDonde   As String

102         respuesta = Reader.ReadBool()
        
            Dim Log As String

104         Log = "Repuesta "

106         If respuesta Then
        
108             Select Case UserList(UserIndex).flags.pregunta

                    Case 1
110                     Log = "Repuesta Afirmativa 1"

                        'Call WriteConsoleMsg(UserIndex, "El usuario desea unirse al grupo.", e_FontTypeNames.FONTTYPE_SUBASTA)
                        ' UserList(UserIndex).Grupo.PropuestaDe = 0
112                     If UserList(UserIndex).Grupo.PropuestaDe <> 0 Then
                
114                         If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Lider <> UserList(UserIndex).Grupo.PropuestaDe Then
116                             Call WriteConsoleMsg(UserIndex, "¡El lider del grupo a cambiado, imposible unirse!", e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
                        
118                             Log = "Repuesta Afirmativa 1-1 "
                        
120                             If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Lider = 0 Then
122                                 Call WriteConsoleMsg(UserIndex, "¡El grupo ya no existe!", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Else
                            
124                                 Log = "Repuesta Afirmativa 1-2 "
                            
126                                 If UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros = 1 Then
128                                     Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe, "36", e_FontTypeNames.FONTTYPE_INFOIAO)
                                        'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "íEl grupo a sido creado!", e_FontTypeNames.FONTTYPE_INFOIAO)
130                                     UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.EnGrupo = True
132                                     Log = "Repuesta Afirmativa 1-3 "

                                    End If
                                
134                                 Log = "Repuesta Afirmativa 1-4"
136                                 UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros = UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros + 1
138                                 UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Miembros(UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros) = UserIndex
140                                 UserList(UserIndex).Grupo.EnGrupo = True
                                
                                    Dim Index As Byte
                                
142                                 Log = "Repuesta Afirmativa 1-5 "
                                
144                                 For Index = 2 To UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros - 1
146                                     Call WriteLocaleMsg(UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Miembros(Index), "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).Name)
                                
148                                 Next Index
                                
150                                 Log = "Repuesta Afirmativa 1-6 "
                                    'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "í" & UserList(UserIndex).name & " a sido añadido al grupo!", e_FontTypeNames.FONTTYPE_INFOIAO)
152                                 Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe, "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).Name)
                                
154                                 Call WriteConsoleMsg(UserIndex, "¡Has sido añadido al grupo!", e_FontTypeNames.FONTTYPE_INFOIAO)
                                
156                                 Log = "Repuesta Afirmativa 1-7 "
                                
158                                 Call RefreshCharStatus(UserList(UserIndex).Grupo.PropuestaDe)
160                                 Call RefreshCharStatus(UserIndex)
                                 
162                                 Log = "Repuesta Afirmativa 1-8"

164                                 Call CompartirUbicacion(UserIndex)

                                End If

                            End If

                        Else
                    
166                         Call WriteConsoleMsg(UserIndex, "Servidor » Solicitud de grupo invalida, reintente...", e_FontTypeNames.FONTTYPE_SERVER)
                    
                        End If

                        'unirlo
168                 Case 2
170                     Log = "Repuesta Afirmativa 2"
172                     Call WriteConsoleMsg(UserIndex, "¡Ahora sos un ciudadano!", e_FontTypeNames.FONTTYPE_INFOIAO)
174                     Call VolverCiudadano(UserIndex)
                    
176                 Case 3
178                     Log = "Repuesta Afirmativa 3"
                    
180                     UserList(UserIndex).Hogar = UserList(UserIndex).PosibleHogar

182                     Select Case UserList(UserIndex).Hogar

                            Case e_Ciudad.cUllathorpe
184                             DeDonde = "Ullathorpe"
                            
186                         Case e_Ciudad.cNix
188                             DeDonde = "Nix"
                
190                         Case e_Ciudad.cBanderbill
192                             DeDonde = "Banderbill"
                        
194                         Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
196                             DeDonde = "Lindos"
                            
198                         Case e_Ciudad.cArghal
200                             DeDonde = " Arghal"
                            
202                         Case e_Ciudad.cArkhein
204                             DeDonde = " Arkhein"
                            
206                         Case Else
208                             DeDonde = "Ullathorpe"

                        End Select
                    
210                     If UserList(UserIndex).flags.TargetNPC <> 0 Then
                    
212                         Call WriteChatOverHead(UserIndex, "¡Gracias " & UserList(UserIndex).Name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Else
214                         Call WriteConsoleMsg(UserIndex, "¡Gracias " & UserList(UserIndex).Name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                    
216                 Case 4
218                     Log = "Repuesta Afirmativa 4"
                
220                     If UserList(UserIndex).flags.TargetUser <> 0 Then
                
222                         UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
224                         UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
226                         UserList(UserIndex).ComUsu.cant = 0
228                         UserList(UserIndex).ComUsu.Objeto = 0
230                         UserList(UserIndex).ComUsu.Acepto = False
                    
                            'Rutina para comerciar con otro usuario
232                         Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)

                        Else
234                         Call WriteConsoleMsg(UserIndex, "Servidor » Solicitud de comercio invalida, reintente...", e_FontTypeNames.FONTTYPE_SERVER)
                
                        End If
                
236                 Case 5
238                     Log = "Repuesta Afirmativa 5"
                
240                     If MapInfo(.Pos.Map).Newbie Then
242                         Call WarpToLegalPos(UserIndex, 140, 53, 58)
244                         .Counters.TimerBarra = 5
246                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, e_ParticulasIndex.Resucitar, .Counters.TimerBarra, False))
248                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, .Counters.TimerBarra, e_AccionBarra.Resucitar))
250                         UserList(UserIndex).Accion.AccionPendiente = True
252                         UserList(UserIndex).Accion.Particula = e_ParticulasIndex.Resucitar
254                         UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.Resucitar
    
256                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", .Pos.x, .Pos.Y))
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", e_FontTypeNames.FONTTYPE_INFO)
258                         Call WriteLocaleMsg(UserIndex, "82", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Else
260                         Call WriteConsoleMsg(UserIndex, "Ya no te encuentras en un mapa newbie.", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
262                 Case Else
264                     Call WriteConsoleMsg(UserIndex, "No tienes preguntas pendientes.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    
                End Select
        
            Else
266             Log = "Repuesta negativa"
        
268             Select Case UserList(UserIndex).flags.pregunta

                    Case 1
270                     Log = "Repuesta negativa 1"

272                     If UserList(UserIndex).Grupo.PropuestaDe <> 0 Then
274                         Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "El usuario no esta interesado en formar parte del grupo.", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If

276                     UserList(UserIndex).Grupo.PropuestaDe = 0
278                     Call WriteConsoleMsg(UserIndex, "Has rechazado la propuesta.", e_FontTypeNames.FONTTYPE_INFOIAO)
                
280                 Case 2
282                     Log = "Repuesta negativa 2"
284                     Call WriteConsoleMsg(UserIndex, "¡Continuas siendo neutral!", e_FontTypeNames.FONTTYPE_INFOIAO)
286                     Call VolverCriminal(UserIndex)

288                 Case 3
290                     Log = "Repuesta negativa 3"
                    
292                     Select Case UserList(UserIndex).PosibleHogar

                            Case e_Ciudad.cUllathorpe
294                             DeDonde = "Ullathorpe"
                            
296                         Case e_Ciudad.cNix
298                             DeDonde = "Nix"
                
300                         Case e_Ciudad.cBanderbill
302                             DeDonde = "Banderbill"
                        
304                         Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
306                             DeDonde = "Lindos"
                            
308                         Case e_Ciudad.cArghal
310                             DeDonde = " Arghal"
                            
312                         Case e_Ciudad.cArkhein
314                             DeDonde = " Arkhein"
                            
316                         Case Else
318                             DeDonde = "Ullathorpe"

                        End Select
                    
320                     If UserList(UserIndex).flags.TargetNPC <> 0 Then
322                         Call WriteChatOverHead(UserIndex, "¡No hay problema " & UserList(UserIndex).Name & "! Sos bienvenido en " & DeDonde & " cuando gustes.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)

                        End If

324                     UserList(UserIndex).PosibleHogar = UserList(UserIndex).Hogar
                    
326                 Case 4
328                     Log = "Repuesta negativa 4"
                    
330                     If UserList(UserIndex).flags.TargetUser <> 0 Then
332                         Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "El usuario no desea comerciar en este momento.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

334                 Case 5
336                     Log = "Repuesta negativa 5"
                        'No hago nada. dijo que no lo resucite
                        
338                 Case Else
340                     Call WriteConsoleMsg(UserIndex, "No tienes preguntas pendientes.", e_FontTypeNames.FONTTYPE_INFOIAO)

                End Select
            
            End If

        End With
    
        Exit Sub
    
errHandler:
342     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResponderPregunta", Erl)
344

End Sub

Private Sub HandleRequestGrupo(ByVal UserIndex As Integer)

        On Error GoTo hErr

        'Author: Pablo Mercavides

100     Call WriteDatosGrupo(UserIndex)
    
        Exit Sub
    
hErr:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGrupo", Erl)
104

End Sub

Private Sub HandleAbandonarGrupo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleAbandonarGrupo_Err

100     With UserList(UserIndex)

        
102         Call Reader.ReadInt16
        
104         If UserList(UserIndex).Grupo.Lider = UserIndex Then
            
106             Call FinalizarGrupo(UserIndex)

                Dim i As Byte
            
108             For i = 2 To UserList(UserIndex).Grupo.CantidadMiembros
110                 Call WriteUbicacion(UserIndex, i, 0)
112             Next i

114             UserList(UserIndex).Grupo.CantidadMiembros = 0
116             UserList(UserIndex).Grupo.EnGrupo = False
118             UserList(UserIndex).Grupo.Lider = 0
120             UserList(UserIndex).Grupo.PropuestaDe = 0
122             Call WriteConsoleMsg(UserIndex, "Has disuelto el grupo.", e_FontTypeNames.FONTTYPE_INFOIAO)
124             Call RefreshCharStatus(UserIndex)
            
            Else
126             Call SalirDeGrupo(UserIndex)

            End If

        End With
        
        Exit Sub

HandleAbandonarGrupo_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAbandonarGrupo", Erl)
130
    
End Sub

Private Sub HandleHecharDeGrupo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleHecharDeGrupo_Err

100     With UserList(UserIndex)

            Dim Indice As Byte

102         Indice = Reader.ReadInt8()
        
104         Call EcharMiembro(UserIndex, Indice)

        End With
        
        Exit Sub

HandleHecharDeGrupo_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHecharDeGrupo", Erl)
108
    
End Sub

Private Sub HandleMacroPos(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleMacroPos_Err

100     With UserList(UserIndex)

102         .ChatCombate = Reader.ReadInt8()
104         .ChatGlobal = Reader.ReadInt8()

        End With
        
        Exit Sub

HandleMacroPos_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMacroPos", Erl)
108
    
End Sub

Private Sub HandleSubastaInfo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleSubastaInfo_Err

100     With UserList(UserIndex)

102         If Subasta.HaySubastaActiva Then

104             Call WriteConsoleMsg(UserIndex, "Subastador: " & Subasta.Subastador, e_FontTypeNames.FONTTYPE_SUBASTA)
106             Call WriteConsoleMsg(UserIndex, "Objeto: " & ObjData(Subasta.ObjSubastado).Name & " (" & Subasta.ObjSubastadoCantidad & ")", e_FontTypeNames.FONTTYPE_SUBASTA)

108             If Subasta.HuboOferta Then
110                 Call WriteConsoleMsg(UserIndex, "Mejor oferta: " & PonerPuntos(Subasta.MejorOferta) & " monedas de oro por " & Subasta.Comprador & ".", e_FontTypeNames.FONTTYPE_SUBASTA)
112                 Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & PonerPuntos(Subasta.MejorOferta + 100), e_FontTypeNames.FONTTYPE_SUBASTA)
                Else
114                 Call WriteConsoleMsg(UserIndex, "Oferta inicial: " & PonerPuntos(Subasta.OfertaInicial) & " monedas de oro.", e_FontTypeNames.FONTTYPE_SUBASTA)
116                 Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & PonerPuntos(Subasta.OfertaInicial + 100), e_FontTypeNames.FONTTYPE_SUBASTA)

                End If

118             Call WriteConsoleMsg(UserIndex, "Tiempo Restante de subasta:  " & SumarTiempo(Subasta.TiempoRestanteSubasta), e_FontTypeNames.FONTTYPE_SUBASTA)
            
            Else
120             Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta activa en este momento.", e_FontTypeNames.FONTTYPE_SUBASTA)

            End If

        End With
        
        Exit Sub

HandleSubastaInfo_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSubastaInfo", Erl)
124
End Sub

Private Sub HandleCancelarExit(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleCancelarExit_Err

100     Call CancelExit(UserIndex)
        
        Exit Sub

HandleCancelarExit_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
104
        
End Sub

Private Sub HandleBanCuenta(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Reason   As String
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
108             Call BanearCuenta(UserIndex, UserName, Reason)
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanCuenta", Erl)
114

End Sub

Private Sub HandleUnBanCuenta(ByVal UserIndex As Integer)

        ' /unbancuenta namepj
        ' /unbancuenta email
        '***************************************************

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserNameOEmail As String
102         UserNameOEmail = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
116             If DesbanearCuenta(UserIndex, UserNameOEmail) Then
118                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & .Name & " ha desbaneado la cuenta de " & UserNameOEmail & ".", e_FontTypeNames.FONTTYPE_SERVER))
                Else
                    Call WriteConsoleMsg(UserIndex, "No se ha podido desbanear la cuenta.", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnBanCuenta", Erl)
124

End Sub

Private Sub HandleCerrarCliente(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As Integer
         
102         UserName = Reader.ReadString8()
        
            ' Solo administradores pueden cerrar clientes ajenos
104         If (.flags.Privilegios And e_PlayerType.Admin) Then

106             tUser = NameIndex(UserName)
            
108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " cerro el cliente de " & UserName & ".", e_FontTypeNames.FONTTYPE_INFO))
                    
114                 Call WriteCerrarleCliente(tUser)

116                 Call LogGM(.Name, "Cerro el cliene de:" & UserName)

                End If

            End If

        End With

        Exit Sub

errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCerrarCliente", Erl)
120

End Sub

Private Sub HandleEventoInfo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleEventoInfo_Err

100     With UserList(UserIndex)

102         If EventoActivo Then
104             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", e_FontTypeNames.FONTTYPE_New_Eventos)
            Else
106             Call WriteConsoleMsg(UserIndex, "Eventos> Actualmente no hay ningun evento en curso.", e_FontTypeNames.FONTTYPE_New_Eventos)

            End If
        
            Dim i           As Byte
            Dim encontre    As Boolean
            Dim HoraProximo As Byte
   
108         If Not HoraEvento + 1 >= 24 Then
   
110             For i = HoraEvento + 1 To 23

112                 If Evento(i).Tipo <> 0 Then
114                     encontre = True
116                     HoraProximo = i
                        Exit For

                    End If

118             Next i

            End If
        
120         If encontre = False Then

122             For i = 0 To HoraEvento

124                 If Evento(i).Tipo <> 0 Then
126                     encontre = True
128                     HoraProximo = i
                        Exit For

                    End If

130             Next i

            End If
        
132         If encontre Then
134             Call WriteConsoleMsg(UserIndex, "Eventos> El proximo evento " & DescribirEvento(HoraProximo) & " iniciara a las " & HoraProximo & ":00 horas.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Else
136             Call WriteConsoleMsg(UserIndex, "Eventos> No hay eventos proximos.", e_FontTypeNames.FONTTYPE_New_Eventos)

            End If

        End With
        
        Exit Sub

HandleEventoInfo_Err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
140
End Sub

Private Sub HandleCrearEvento(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Pablo Mercavides
        '***************************************************

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Tipo           As Byte
            Dim Duracion       As Byte
            Dim multiplicacion As Byte
        
102         Tipo = Reader.ReadInt8()
104         Duracion = Reader.ReadInt8()
106         multiplicacion = Reader.ReadInt8()

108         If multiplicacion > 5 Then 'no superar este multiplicador
110             multiplicacion = 2
            End If
        
            '/ dejar solo Administradores
112         If .flags.Privilegios >= e_PlayerType.Admin Then
114             If EventoActivo = False Then
116                 If LenB(Tipo) = 0 Or LenB(Duracion) = 0 Or LenB(multiplicacion) = 0 Then
118                     Call WriteConsoleMsg(UserIndex, "Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.", e_FontTypeNames.FONTTYPE_New_Eventos)
                    Else
                
120                     Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).Name)
                  
                    End If

                Else
122                 Call WriteConsoleMsg(UserIndex, "Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.", e_FontTypeNames.FONTTYPE_New_Eventos)

                End If
            Else
124             Call WriteConsoleMsg(UserIndex, "Servidor » Solo Administradores pueder crear estos eventos.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

errHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
128

End Sub

Private Sub HandleBanTemporal(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

        On Error GoTo errHandler

100     With UserList(UserIndex)
         
            Dim UserName As String
            Dim Reason   As String
            Dim dias     As Byte
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         dias = Reader.ReadInt8()
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
110             Call Admin.BanTemporal(UserName, dias, Reason, UserList(UserIndex).Name)
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
116

End Sub

Private Sub HandleCompletarViaje(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Destino As Byte

            Dim costo   As Long

102         Destino = Reader.ReadInt8()
104         costo = Reader.ReadInt32()

            ' WyroX: WTF el costo lo decide el cliente... Desactivo....
            Exit Sub

106         If costo <= 0 Then Exit Sub

            Dim DeDonde As t_CityWorldPos

108         If UserList(UserIndex).Stats.GLD < costo Then
110             Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", e_FontTypeNames.FONTTYPE_INFO)
            
            Else

112             Select Case Destino

                    Case e_Ciudad.cUllathorpe
114                     DeDonde = CityUllathorpe
                        
116                 Case e_Ciudad.cNix
118                     DeDonde = CityNix
            
120                 Case e_Ciudad.cBanderbill
122                     DeDonde = CityBanderbill
                    
124                 Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
126                     DeDonde = CityLindos
                        
128                 Case e_Ciudad.cArghal
130                     DeDonde = CityArghal
                        
132                 Case e_Ciudad.cArkhein
134                     DeDonde = CityArkhein
                        
136                 Case Else
138                     DeDonde = CityUllathorpe

                End Select
        
140             If DeDonde.NecesitaNave > 0 Then
142                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Navegacion) < 80 Then
                        Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", e_FontTypeNames.FONTTYPE_INFO)
144                     Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", e_FontTypeNames.FONTTYPE_WARNING)
                    Else

146                     If UserList(UserIndex).flags.TargetNPC <> 0 Then
148                         If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
150                             Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                            End If

                        End If

152                     Call WarpToLegalPos(UserIndex, DeDonde.MapaViaje, DeDonde.ViajeX, DeDonde.ViajeY, True)
154                     Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", e_FontTypeNames.FONTTYPE_WARNING)
156                     UserList(UserIndex).Stats.MinAGU = 0
158                     UserList(UserIndex).Stats.MinHam = 0
160                     UserList(UserIndex).flags.Sed = 1
162                     UserList(UserIndex).flags.Hambre = 1
                    
164                     UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
166                     Call WriteUpdateHungerAndThirst(UserIndex)
168                     Call WriteUpdateUserStats(UserIndex)

                    End If

                Else
            
                    Dim Map As Integer

                    Dim x   As Byte

                    Dim Y   As Byte
            
170                 Map = DeDonde.MapaViaje
172                 x = DeDonde.ViajeX
174                 Y = DeDonde.ViajeY

176                 If UserList(UserIndex).flags.TargetNPC <> 0 Then
178                     If NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
180                         Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                        End If

                    End If
                
182                 Call WarpUserChar(UserIndex, Map, x, Y, True)
184                 Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", e_FontTypeNames.FONTTYPE_WARNING)
186                 UserList(UserIndex).Stats.MinAGU = 0
188                 UserList(UserIndex).Stats.MinHam = 0
190                 UserList(UserIndex).flags.Sed = 1
192                 UserList(UserIndex).flags.Hambre = 1
                
194                 UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
196                 Call WriteUpdateHungerAndThirst(UserIndex)
198                 Call WriteUpdateUserStats(UserIndex)
        
                End If

            End If

        End With
    
        Exit Sub

errHandler:
200     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCompletarViaje", Erl)
202

End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuest_Err

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer
        Dim tmpByte  As Byte

100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
104     If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
106         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
108     If NpcList(NpcIndex).NumQuest = 0 Then
110         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageChatOverHead("No tengo ninguna misión para ti.", NpcList(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
    
112     Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", NpcList(NpcIndex).Char.CharIndex, vbWhite))

        Exit Sub

HandleQuest_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuest", Erl)
116
        
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
        Dim QuestNpc As Boolean
        
100     Indice = Reader.ReadInt8
        QuestNpc = Reader.ReadBool
102     NpcIndex = UserList(UserIndex).flags.TargetNPC
104     If NpcIndex = 0 And QuestNpc = 1 Then Exit Sub

        If QuestNpc Then
106         If Indice <= 0 Or Indice > UBound(NpcList(NpcIndex).QuestNumber) Then Exit Sub
        Else
             If Indice <= 0 Then Exit Sub
        End If
        
    If QuestNpc Then
        
        'Esta el personaje en la distancia correcta?
108     If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
110         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
112     If TieneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
114         Call WriteConsoleMsg(UserIndex, "La quest ya esta en curso.", e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If
        
        'El personaje completo la quest que requiere?
116     If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest > 0 Then
118         If Not UserDoneQuest(UserIndex, QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest) Then
120             Call WriteChatOverHead(UserIndex, "Debes completar la quest " & QuestList(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest).nombre & " para emprender esta misión.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                Exit Sub

            End If

        End If

        'El personaje tiene suficiente nivel?
122     If UserList(UserIndex).Stats.ELV < QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel Then
124         Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel & " para emprender esta misión.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub

        End If
        
        'La quest no es repetible?
        If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).Repetible = 0 Then
            'El personaje ya hizo la quest?
126         If UserDoneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
128             Call WriteChatOverHead(UserIndex, "QUESTNEXT*" & NpcList(NpcIndex).QuestNumber(Indice), NpcList(NpcIndex).Char.CharIndex, vbYellow)
                Exit Sub
    
            End If
        End If
    
130     QuestSlot = FreeQuestSlot(UserIndex)

132     If QuestSlot = 0 Then
134         Call WriteChatOverHead(UserIndex, "Debes completar las misiones en curso para poder aceptar más misiones.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub

        End If
    
        'Agregamos la quest.
136     With UserList(UserIndex).QuestStats.Quests(QuestSlot)
138         .QuestIndex = UserList(UserIndex).flags.QuestNumber
        
140         If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
142         If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
            UserList(UserIndex).flags.ModificoQuests = True
            
144         Call WriteConsoleMsg(UserIndex, "Has aceptado la misión " & Chr(34) & QuestList(.QuestIndex).nombre & Chr(34) & ".", e_FontTypeNames.FONTTYPE_INFOIAO)
146
            If (FinishQuestCheck(UserIndex, .QuestIndex, QuestSlot)) Then
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 3)
            Else
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
            End If
            
        End With
        
            
    Else
         If TieneQuest(UserIndex, UserList(UserIndex).flags.QuestNumber) Then
             Call WriteConsoleMsg(UserIndex, "La quest ya esta en curso.", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
            
            'El personaje tiene suficiente nivel?
         If UserList(UserIndex).Stats.ELV < QuestList(UserList(UserIndex).flags.QuestNumber).RequiredLevel Then
                Call WriteConsoleMsg(UserIndex, "Debes ser por lo menos nivel " & QuestList(UserList(UserIndex).flags.QuestNumber).RequiredLevel & " para emprender esta misión.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
         QuestSlot = FreeQuestSlot(UserIndex)
    
         If QuestSlot = 0 Then
                Call WriteConsoleMsg(UserIndex, "Debes completar las misiones en curso para poder aceptar más misiones.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Agregamos la quest.
            With UserList(UserIndex).QuestStats.Quests(QuestSlot)
             .QuestIndex = UserList(UserIndex).flags.QuestNumber
    
            If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
            If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
                UserList(UserIndex).flags.ModificoQuests = True
            
            
            Call WriteConsoleMsg(UserIndex, "Has aceptado la misión " & Chr(34) & QuestList(.QuestIndex).nombre & Chr(34) & ".", e_FontTypeNames.FONTTYPE_INFOIAO)
    
            If (FinishQuestCheck(UserIndex, .QuestIndex, QuestSlot)) Then
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 3)
            Else
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
            End If
            Call QuitarUserInvItem(UserIndex, UserList(UserIndex).flags.QuestItemSlot, 1)
            Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.QuestItemSlot)
        End With
    End If
        Exit Sub

HandleQuestAccept_Err:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAccept", Erl)
150
        
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestDetailsRequest_Err

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestInfoRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim QuestSlot As Byte

100     QuestSlot = Reader.ReadInt8
    
102     Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
        
        Exit Sub

HandleQuestDetailsRequest_Err:
104     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestDetailsRequest", Erl)
106
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestAbandon.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        On Error GoTo HandleQuestAbandon_Err
        
        With UserList(UserIndex)
        
            Dim Slot As Byte
            Slot = Reader.ReadInt8
            
            With .QuestStats.Quests(Slot)
                ' Le quitamos los objetos de quest que no puede tirar
                If QuestList(.QuestIndex).RequiredOBJs Then
                
                    Dim ObjIndex As Integer, i As Integer
                    
                    For i = 1 To QuestList(.QuestIndex).RequiredOBJs
                        ObjIndex = QuestList(.QuestIndex).RequiredOBJ(i).ObjIndex
                        
                        If ObjData(ObjIndex).Intirable = 1 And ObjData(ObjIndex).Instransferible Then
                        
                            ' Revisamos que ninguna otra quest que tenga activa le pida el mismo item
                            Dim q As Integer, j As Byte, K As Byte, QuitarItem As Boolean

                            QuitarItem = True
                            
                            For j = 1 To MAXUSERQUESTS
                                q = UserList(UserIndex).QuestStats.Quests(j).QuestIndex
                                If q <> 0 And q <> .QuestIndex Then
                                    For K = 1 To QuestList(q).RequiredOBJs
                                        If QuestList(q).RequiredOBJ(K).ObjIndex = ObjIndex Then
                                            QuitarItem = False
                                            Exit For
                                        End If
                                    Next
                                End If
                                If Not QuitarItem Then Exit For
                            Next
                            
                            If QuitarItem Then
                                Call QuitarObjetos(ObjIndex, MAX_INVENTORY_OBJS, UserIndex)
                            End If
                        End If
                    Next i
                
                End If
            End With
    
            'Borramos la quest.
100         Call CleanQuestSlot(UserIndex, Slot)
        
            'Ordenamos la lista de quests del usuario.
102         Call ArrangeUserQuests(UserIndex)
        
            'Enviamos la lista de quests actualizada.
104         Call WriteQuestListSend(UserIndex)

        End With
        
        Exit Sub

HandleQuestAbandon_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAbandon", Erl)
108
        
End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestListRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error GoTo HandleQuestListRequest_Err

100     Call WriteQuestListSend(UserIndex)
        
        Exit Sub

HandleQuestListRequest_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestListRequest", Erl)
104
        
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

        On Error GoTo errHandler

        Dim Map   As Integer
        Dim x     As Byte
        Dim Y     As Byte
        Dim Index As Long
    
100     With UserList(UserIndex)

102         Map = Reader.ReadInt16()
104         x = Reader.ReadInt8()
106         Y = Reader.ReadInt8()
        
            ' User Admin?
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            ' Valid pos?
112         If Not InMapBounds(Map, x, Y) Then
114             Call WriteConsoleMsg(UserIndex, "Posicion invalida.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            ' Choose pretorian clan index
116         If Map = MAPA_PRETORIANO Then
118             Index = e_PretorianType.Default ' Default clan
            
            Else
120             Index = e_PretorianType.Custom ' Custom Clan

            End If
            
            ' Is already active any clan?
122         If Not ClanPretoriano(Index).Active Then
            
124             If Not ClanPretoriano(Index).SpawnClan(Map, x, Y, Index) Then
126                 Call WriteConsoleMsg(UserIndex, "La posicion no es apropiada para crear el clan", e_FontTypeNames.FONTTYPE_INFO)

                End If
        
            Else
128             Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", e_FontTypeNames.FONTTYPE_INFO)

            End If
    
        End With

        Exit Sub

errHandler:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreatePretorianClan", Erl)
132
    
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

        On Error GoTo errHandler
    
        Dim Map   As Integer
        Dim Index As Long
    
100     With UserList(UserIndex)

102         Map = Reader.ReadInt16()
        
            ' User Admin?
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            ' Valid map?
108         If Map < 1 Or Map > NumMaps Then
110             Call WriteConsoleMsg(UserIndex, "Mapa invalido.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Los sacamos correctamente.
112         Call EliminarPretorianos(Map)
    
        End With

        Exit Sub

errHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreatePretorianClan", Erl)
116

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
    
100     With UserList(UserIndex)
 
            Dim Nick As String
102         Nick = Reader.ReadString8

            ' Comando exclusivo para gms
104         If Not EsGM(UserIndex) Then Exit Sub
        
106         If Len(Nick) <> 0 Then
108             UserConsulta = NameIndex(Nick)
            
                'Se asegura que el target exista
110             If UserConsulta <= 0 Then
112                 Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
    
                End If
            
            Else
        
114             UserConsulta = .flags.TargetUser
            
                'Se asegura que el target exista
116             If UserConsulta <= 0 Then
118                 Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
    
                End If
            
            End If

            ' No podes ponerte a vos mismo en modo consulta.
120         If UserConsulta = UserIndex Then Exit Sub
        
            ' No podes estra en consulta con otro gm
122         If EsGM(UserConsulta) Then
124             Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            ' Si ya estaba en consulta, termina la consulta
126         If UserList(UserConsulta).flags.EnConsulta Then
128             Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserList(UserConsulta).Name & ".", e_FontTypeNames.FONTTYPE_INFOBOLD)
130             Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", e_FontTypeNames.FONTTYPE_INFOBOLD)
            
132             Call LogGM(.Name, "Termino consulta con " & UserList(UserConsulta).Name)
            
134             UserList(UserConsulta).flags.EnConsulta = False
        
                ' Sino la inicia
            Else
        
136             Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserList(UserConsulta).Name & ".", e_FontTypeNames.FONTTYPE_INFOBOLD)
138             Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", e_FontTypeNames.FONTTYPE_INFOBOLD)
            
140             Call LogGM(.Name, "Inicio consulta con " & UserList(UserConsulta).Name)
            
142             With UserList(UserConsulta)

144                 If Not EstaPCarea(UserIndex, UserConsulta) Then
                        Dim x As Byte
                        Dim Y As Byte
                        
146                     x = .Pos.x
148                     Y = .Pos.Y
150                     Call FindLegalPos(UserIndex, .Pos.Map, x, Y)
152                     Call WarpUserChar(UserIndex, .Pos.Map, x, Y, True)
                        
                    End If
            
154                 If UserList(UserIndex).flags.AdminInvisible = 1 Then
156                     Call DoAdminInvisible(UserIndex)

                    End If

158                 .flags.EnConsulta = True
                
                    ' Pierde invi u ocu
160                 If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                
162                     .flags.Oculto = 0
164                     .flags.invisible = 0
166                     .Counters.TiempoOculto = 0
168                     .Counters.Invisibilidad = 0
                    
170                     If UserList(UserConsulta).flags.Navegando = 0 Then
                            
172                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                        End If

                    End If

                End With

            End If
        
174         Call SetModoConsulta(UserConsulta)

        End With
    
        Exit Sub
    
errHandler:
176     Call TraceError(Err.Number, Err.Description, "Protocol.HandleConsulta", Erl)
178

End Sub

Private Sub HandleGetMapInfo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then
            
                Dim Response As String
            
104             Response = "[Info de mapa " & .Pos.Map & "]" & vbNewLine
106             Response = Response & "Nombre = " & MapInfo(.Pos.Map).map_name & vbNewLine
108             Response = Response & "Seguro = " & MapInfo(.Pos.Map).Seguro & vbNewLine
110             Response = Response & "Newbie = " & MapInfo(.Pos.Map).Newbie & vbNewLine
112             Response = Response & "Nivel = " & MapInfo(.Pos.Map).MinLevel & "/" & MapInfo(.Pos.Map).MaxLevel & vbNewLine
114             Response = Response & "SinInviOcul = " & MapInfo(.Pos.Map).SinInviOcul & vbNewLine
116             Response = Response & "SinMagia = " & MapInfo(.Pos.Map).SinMagia & vbNewLine
118             Response = Response & "SoloClanes = " & MapInfo(.Pos.Map).SoloClanes & vbNewLine
120             Response = Response & "NoPKs = " & MapInfo(.Pos.Map).NoPKs & vbNewLine
122             Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
124             Response = Response & "Salida = " & MapInfo(.Pos.Map).Salida.Map & "-" & MapInfo(.Pos.Map).Salida.x & "-" & MapInfo(.Pos.Map).Salida.Y & vbNewLine
126             Response = Response & "Terreno = " & MapInfo(.Pos.Map).terrain & vbNewLine
128             Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
130             Response = Response & "Zona = " & MapInfo(.Pos.Map).zone & vbNewLine
            
132             Call WriteConsoleMsg(UserIndex, Response, e_FontTypeNames.FONTTYPE_INFO)
        
            End If
    
        End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Name As String
102         Name = Reader.ReadString8()

104         If LenB(Name) = 0 Then Exit Sub

106         If EsGmChar(Name) Then
108             Call WriteConsoleMsg(UserIndex, "No podés denunciar a un administrador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Dim tUser As Integer
110         tUser = NameIndex(Name)
        
112         If tUser <= 0 Then
114             Call WriteConsoleMsg(UserIndex, "El usuario no está online.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Dim Denuncia As String, HayChat As Boolean
116         Denuncia = "[Últimos mensajes de " & UserList(tUser).Name & "]" & vbNewLine
        
            Dim i As Integer

118         For i = 1 To UBound(UserList(tUser).flags.ChatHistory)

120             If LenB(UserList(tUser).flags.ChatHistory(i)) <> 0 Then
122                 Denuncia = Denuncia & UserList(tUser).flags.ChatHistory(i) & vbNewLine
124                 HayChat = True

                End If

            Next
        
126         If Not HayChat Then
128             Call WriteConsoleMsg(UserIndex, "El usuario no ha escrito nada. Recordá que las denuncias inválidas pueden ser motivo de advertencia.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

136         Call Ayuda.Push(.Name, Denuncia, "Denuncia a " & UserList(tUser).Name)
138         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido una nueva denuncia de parte de " & .Name & ".", e_FontTypeNames.FONTTYPE_SERVER))

140         Call WriteConsoleMsg(UserIndex, "Tu denuncia fue recibida por el equipo de soporte.", e_FontTypeNames.FONTTYPE_INFOIAO)

142         Call LogConsulta(.Name & " (Denuncia a " & UserList(tUser).Name & ")" & vbNewLine & Denuncia)

        End With
    
        Exit Sub
    
        Exit Sub

errHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
146

End Sub

Private Sub HandleSeguroResu(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         .flags.SeguroResu = Not .flags.SeguroResu
        
104         Call WriteSeguroResu(UserIndex, .flags.SeguroResu)
    
        End With

End Sub

Private Sub HandleCuentaExtractItem(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCuentaExtractItem_Err

        '***************************************************
        'Author: Ladder
        'Last Modification: 22/11/21
        'Retirar item de cuenta
        '***************************************************

100     With UserList(UserIndex)

            Dim Slot        As Byte

            Dim slotdestino As Byte

            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
        
106         slotdestino = Reader.ReadInt8()
        
108         If .flags.Muerto = 1 Then
110             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If .flags.TargetNPC < 1 Then Exit Sub
        
114         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then
                Exit Sub

            End If
        
            'acá va el guardado en memoria
        
            'User retira el item del slot
            'Call UserRetiraItem(UserIndex, slot, Amount, slotdestino)

        End With
        
        Exit Sub

HandleCuentaExtractItem_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaExtractItem", Erl)
118
        
End Sub

Private Sub HandleCuentaDeposit(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCuentaDeposit_Err

        '***************************************************
        'Author: Ladder
        'Last Modification: 22/11/21
        'Depositar item en cuenta
        '***************************************************
    
100     With UserList(UserIndex)

            Dim Slot        As Byte

            Dim slotdestino As Byte

            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce...
108         If .flags.Muerto = 1 Then
110             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
112         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
114         If NpcList(.flags.TargetNPC).NPCtype <> e_NPCType.Banquero Then
                Exit Sub

            End If
            
116         If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 10 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'acá va el guardado en memoria
            
            'User deposita el item del slot rdata
            'Call UserDepositaItem(UserIndex, slot, Amount, slotdestino)

        End With
        
        Exit Sub

HandleCuentaDeposit_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaDeposit", Erl)
122
        
End Sub

Private Sub HandleCommerceSendChatMessage(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)

            Dim chatMessage As String
        
102         chatMessage = "[" & UserList(UserIndex).Name & "] " & Reader.ReadString8
        
            'El mensaje se lo envío al destino
104         Call WriteCommerceRecieveChatMessage(UserList(UserIndex).ComUsu.DestUsu, chatMessage)
        
            'y tambien a mi mismo
106         Call WriteCommerceRecieveChatMessage(UserIndex, chatMessage)

        End With
    
        Exit Sub
    
errHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSendChatMessage", Erl)
110
    
End Sub

Private Sub HandleLogMacroClickHechizo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("AntiCheat> El usuario " & .Name & " se le cerró el cliente por posible uso de macro de hechizos", e_FontTypeNames.FONTTYPE_INFO))
104         Call LogHackAttemp("Usuario: " & .Name & "   " & "Ip: " & .IP & " Posible uso de macro de hechizos.")

        End With

End Sub

Private Sub HandleCreateEvent(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     With UserList(UserIndex)

            Dim Name As String
102         Name = Reader.ReadString8()

104         If LenB(Name) = 0 Then Exit Sub
    
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
110         Select Case UCase$(Name)

                Case "INVASION BANDER"
112                 Call IniciarEvento(TipoEvento.Invasion, 1)
114                 Call LogGM(.Name, "Forzó el evento Invasión en Banderbille.")
                
116             Case "INVASION CARCEL"
118                 Call IniciarEvento(TipoEvento.Invasion, 2)
120                 Call LogGM(.Name, "Forzó el evento Invasión en Carcel.")

122             Case Else
124                 Call WriteConsoleMsg(UserIndex, "No existe el evento """ & Name & """.", e_FontTypeNames.FONTTYPE_INFO)

            End Select

        End With
    
        Exit Sub

errHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateEvent", Erl)
128
        
End Sub

Private Sub HandleHome(ByVal UserIndex As Integer)
        
        On Error GoTo HandleHome_Err
    
        

        '***************************************************
        'Author: Budi
        'Creation Date: 06/01/2010
        'Last Modification: 05/06/10
        'Pato - 05/06/10: Add the UCase$ to prevent problems.
        '***************************************************
    
100     With UserList(UserIndex)

104         If .flags.Muerto = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
                
            'Si el mapa tiene alguna restriccion (newbie, dungeon, etc...), no lo dejamos viajar.
108         If MapInfo(.Pos.Map).zone = "NEWBIE" Or MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = CARCEL Then
110             Call WriteConsoleMsg(UserIndex, "No pueder viajar a tu hogar desde este mapa.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            
            End If
        
            'Si es un mapa comun y no esta en cana
112         If .Counters.Pena <> 0 Then
114             Call WriteConsoleMsg(UserIndex, "No puedes usar este comando en prisión.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
116         If .flags.EnReto Then
118             Call WriteConsoleMsg(UserIndex, "No podés regresar desde un reto. Usa /ABANDONAR para admitir la derrota y volver a la ciudad.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If

120         If .flags.Traveling = 0 Then
            
122             If .Pos.Map <> Ciudades(.Hogar).Map Then
124                 Call goHome(UserIndex)
                
                Else
126                 Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", e_FontTypeNames.FONTTYPE_INFO)

                End If

            Else

128             .flags.Traveling = 0
130             .Counters.goHome = 0
            
132             Call WriteConsoleMsg(UserIndex, "Ya hay un viaje en curso.", e_FontTypeNames.FONTTYPE_INFO)
            
            End If
        
        End With

        
        Exit Sub

HandleHome_Err:
134     Call TraceError(Err.Number, Err.Description, "Hogar.HandleHome", Erl)

        
End Sub

Private Sub HandleAddItemCrafting(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim InvSlot As Byte, CraftSlot As Byte
102         InvSlot = Reader.ReadInt8
104         CraftSlot = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub
        
108         If InvSlot < 1 Or InvSlot > .CurrentInventorySlots Then Exit Sub

110         If .Invent.Object(InvSlot).ObjIndex = 0 Then Exit Sub

112         If CraftSlot < 1 Then
114             For CraftSlot = 1 To MAX_SLOTS_CRAFTEO
116                 If .CraftInventory(CraftSlot) = 0 Then
                        Exit For
                    End If
                Next
            End If

118         If CraftSlot > MAX_SLOTS_CRAFTEO Then
                Exit Sub
            End If

120         If .CraftInventory(CraftSlot) <> 0 Then Exit Sub

122         .CraftInventory(CraftSlot) = .Invent.Object(InvSlot).ObjIndex
    
124         Call QuitarUserInvItem(UserIndex, InvSlot, 1)
126         Call UpdateUserInv(False, UserIndex, InvSlot)

128         Call WriteCraftingItem(UserIndex, CraftSlot, .CraftInventory(CraftSlot))

            Dim Result As clsCrafteo
130         Set Result = CheckCraftingResult(UserIndex)
        
132         If Not Result Is .CraftResult Then
134             Set .CraftResult = Result
136             If Not .CraftResult Is Nothing Then
138                 Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.Precio)
                Else
140                 Call WriteCraftingResult(UserIndex, 0)
                End If
            End If

        End With
    
        Exit Sub

errHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddItemCrafting", Erl)
144
End Sub

Private Sub HandleRemoveItemCrafting(ByVal UserIndex As Integer)
    
        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim InvSlot As Byte, CraftSlot As Byte
102         CraftSlot = Reader.ReadInt8
104         InvSlot = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub

108         If CraftSlot < 1 Or CraftSlot > MAX_SLOTS_CRAFTEO Then Exit Sub

110         If .CraftInventory(CraftSlot) = 0 Then Exit Sub

112         If InvSlot < 1 Then
                Dim TmpObj As t_Obj
114             TmpObj.ObjIndex = .CraftInventory(CraftSlot)
116             TmpObj.amount = 1
             
118             If Not MeterItemEnInventario(UserIndex, TmpObj) Then Exit Sub

120         ElseIf InvSlot <= .CurrentInventorySlots Then
122             If .Invent.Object(InvSlot).ObjIndex = 0 Then
124                 .Invent.Object(InvSlot).ObjIndex = .CraftInventory(CraftSlot)
            
126             ElseIf .Invent.Object(InvSlot).ObjIndex <> .CraftInventory(CraftSlot) Then
                    Exit Sub
                End If

128             .Invent.Object(InvSlot).amount = .Invent.Object(InvSlot).amount + 1
130             Call UpdateUserInv(False, UserIndex, InvSlot)
            End If

132         .CraftInventory(CraftSlot) = 0
134         Call WriteCraftingItem(UserIndex, CraftSlot, 0)
        
            Dim Result As clsCrafteo
136         Set Result = CheckCraftingResult(UserIndex)
        
138         If Not Result Is .CraftResult Then
140             Set .CraftResult = Result
142             If Not .CraftResult Is Nothing Then
144                 Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.Precio)
                Else
146                 Call WriteCraftingResult(UserIndex, 0)
                End If
            End If

        End With
    
        Exit Sub
    
errHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveItemCrafting", Erl)
150
End Sub

Private Sub HandleAddCatalyst(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim Slot As Byte
102         Slot = Reader.ReadInt8
        
104         If .flags.Crafteando = 0 Then Exit Sub
        
106         If Slot < 1 Or Slot > .CurrentInventorySlots Then Exit Sub

108         If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        
110         If ObjData(.Invent.Object(Slot).ObjIndex).CatalizadorTipo = 0 Then Exit Sub

112         If .CraftCatalyst.ObjIndex <> 0 Then Exit Sub

114         .CraftCatalyst.ObjIndex = .Invent.Object(Slot).ObjIndex
116         .CraftCatalyst.amount = .Invent.Object(Slot).amount

118         Call QuitarUserInvItem(UserIndex, Slot, MAX_INVENTORY_OBJS)
120         Call UpdateUserInv(False, UserIndex, Slot)

122         If .CraftResult Is Nothing Then
124             Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, 0)
            Else
126             Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
            End If

        End With
    
        Exit Sub
    
errHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddCatalyst", Erl)
130
End Sub

Private Sub HandleRemoveCatalyst(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim Slot As Byte
102         Slot = Reader.ReadInt8
        
104         If .flags.Crafteando = 0 Then Exit Sub

106         If .CraftCatalyst.ObjIndex = 0 Then Exit Sub

108         If Slot < 1 Then
110             If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then Exit Sub

112         ElseIf Slot <= .CurrentInventorySlots Then
114             If .Invent.Object(Slot).ObjIndex = 0 Then
116                 .Invent.Object(Slot).ObjIndex = .CraftCatalyst.ObjIndex

118             ElseIf .Invent.Object(Slot).ObjIndex <> .CraftCatalyst.ObjIndex Then
                    Exit Sub
                End If

120             .Invent.Object(Slot).amount = .Invent.Object(Slot).amount + .CraftCatalyst.amount
122             Call UpdateUserInv(False, UserIndex, Slot)
            End If

124         .CraftCatalyst.ObjIndex = 0
126         .CraftCatalyst.amount = 0
        
128         If .CraftResult Is Nothing Then
130             Call WriteCraftingCatalyst(UserIndex, 0, 0, 0)
            Else
132             Call WriteCraftingCatalyst(UserIndex, 0, 0, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
            End If

        End With
    
        Exit Sub
    
errHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCatalyst", Erl)
136
End Sub

Sub HandleCraftItem(ByVal UserIndex As Integer)

        On Error GoTo errHandler

100     If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub

102     Call DoCraftItem(UserIndex)
    
        Exit Sub

errHandler:
104     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftItem", Erl)
106
End Sub

Private Sub HandleCloseCrafting(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub

102     Call ReturnCraftingItems(UserIndex)
    
104     UserList(UserIndex).flags.Crafteando = 0
    
        Exit Sub
    
errHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCloseCrafting", Erl)
108
End Sub

Private Sub HandleMoveCraftItem(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim Drag As Byte, Drop As Byte
102         Drag = Reader.ReadInt8
104         Drop = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub
        
108         If Drag < 1 Or Drag > MAX_SLOTS_CRAFTEO Then Exit Sub
110         If Drop < 1 Or Drop > MAX_SLOTS_CRAFTEO Then Exit Sub
112         If Drag = Drop Then Exit Sub

114         If .CraftInventory(Drag) = 0 Then Exit Sub
116         If .CraftInventory(Drag) = .CraftInventory(Drop) Then Exit Sub

            Dim aux As Integer
118         aux = .CraftInventory(Drop)
120         .CraftInventory(Drop) = .CraftInventory(Drag)
122         .CraftInventory(Drag) = aux

124         Call WriteCraftingItem(UserIndex, Drag, .CraftInventory(Drag))
126         Call WriteCraftingItem(UserIndex, Drop, .CraftInventory(Drop))

        End With
    
        Exit Sub
    
errHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveCraftItem", Erl)
130
End Sub

Private Sub HandlePetLeaveAll(ByVal UserIndex As Integer)

        On Error GoTo errHandler
    
100     With UserList(UserIndex)
    
            Dim AlmenosUna As Boolean, i As Integer
    
102         For i = 1 To MAXMASCOTAS
104             If .MascotasIndex(i) > 0 Then
106                 If NpcList(.MascotasIndex(i)).flags.NPCActive Then
108                     Call QuitarNPC(.MascotasIndex(i))
110                     AlmenosUna = True
                    End If
                End If
112         Next i
        
114         If AlmenosUna Then
                .flags.ModificoMascotas = True
116             Call WriteConsoleMsg(UserIndex, "Liberaste a tus mascotas.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub
    
errHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeaveAll", Erl)
120
End Sub

Private Sub HandleGuardNoticeResponse(ByVal UserIndex As Integer)
    
        On Error GoTo HandleGuardNoticeResponse_Err:

        Dim Codigo As String: Codigo = Reader.ReadString8()

        Call AOGuard.HandleNoticeResponse(UserIndex, Codigo)
        
        Exit Sub

HandleGuardNoticeResponse_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuardNoticeResponse", Erl)

End Sub

Private Sub HandleGuardResendVerificationCode(ByVal UserIndex As Integer)
        
        On Error GoTo HandleResendVerificationCode_Err:
        
100     Call AOGuard.EnviarCodigo(UserIndex)
        
        Exit Sub

HandleResendVerificationCode_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuardResendVerificationCode", Erl)

End Sub

Private Sub HandleResetChar(ByVal UserIndex As Integer)
        On Error GoTo HandleResetChar_Err:
        
100     Dim Nick As String: Nick = Reader.ReadString8()

        #If DEBUGGING = 1 Then

            If UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin Then
                Dim Index As Integer
                Index = NameIndex(Nick)
                
                If Index <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline o inexistente.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                With UserList(Index)
                    .Stats.ELV = 1
                    .Stats.Exp = 0
                    .Stats.SkillPts = 10
                    
                    Dim i As Integer
                    For i = 1 To NUMSKILLS
                        .Stats.UserSkills(i) = 0
                    Next

                    .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
                    .Stats.MinHp = .Stats.MaxHp

                    .Stats.MaxMAN = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
                    .Stats.MinMAN = .Stats.MaxMAN

                    Dim MiInt As Integer
                    MiInt = RandomNumber(1, .Stats.UserAtributosBackUP(e_Atributos.Agilidad) \ 6)

                    If MiInt = 1 Then MiInt = 2
                
                    .Stats.MaxSta = 20 * MiInt
                    .Stats.MinSta = 20 * MiInt
                
                    .Stats.MaxAGU = 100
                    .Stats.MinAGU = 100
                
                    .Stats.MaxHam = 100
                    .Stats.MinHam = 100
            
                    .Stats.MaxHit = 2
                    .Stats.MinHIT = 1
                    
                    .flags.ModificoSkills = True
                    
                    Call WriteUpdateUserStats(Index)
                    Call WriteLevelUp(Index, .Stats.SkillPts)
                End With
                
                Call WriteConsoleMsg(UserIndex, "Personaje reseteado a nivel 1.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        
        #End If
        
        Exit Sub

HandleResetChar_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetChar", Erl)
End Sub

Private Sub HandleDeleteItem(ByVal UserIndex As Integer)
    On Error GoTo HandleDeleteItem_Err:

    Dim Slot As Byte

    Slot = Reader.ReadInt8()

    With UserList(UserIndex)
        If .Invent.Object(Slot).Equipped = 0 Then
            UserList(UserIndex).Invent.Object(Slot).amount = 0
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Call UpdateUserInv(False, UserIndex, Slot)
            Call WriteConsoleMsg(UserIndex, "Objeto eliminado correctamente.", e_FontTypeNames.fonttype_info)
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes eliminar un objeto estando equipado.", e_FontTypeNames.fonttype_info)
            Exit Sub
        End If
    End With

    Exit Sub

HandleDeleteItem_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDeleteItem", Erl)
End Sub
